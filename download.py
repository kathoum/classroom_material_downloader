#!/usr/bin/env python3
import os
import os.path
from dataclasses import dataclass, field
from collections import Counter
from pathlib import Path
from typing import List
import mimetypes
import requests

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.exceptions import RefreshError
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = [
    'https://www.googleapis.com/auth/classroom.courses.readonly',
    'https://www.googleapis.com/auth/classroom.coursework.me.readonly',
    'https://www.googleapis.com/auth/classroom.courseworkmaterials.readonly',
    'https://www.googleapis.com/auth/drive.readonly',
]
os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'


def get_credentials():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def call_list_api(obj, **kwargs):
    result = obj.list(**kwargs, pageSize=100).execute()
    nextPageToken = result.pop('nextPageToken', None)
    while nextPageToken:
        next_result = obj.list(**kwargs, pageSize=100, pageToken=nextPageToken).execute()
        nextPageToken = next_result.pop('nextPageToken', None)
        for k,v in next_result.items():
            result[k] += v
    return result


@dataclass
class Material:
    ID: str
    title: str
    filename: str = None
    size: int = None
    creationTime: str = None
    mimeType: str = None
    exportLinks: [str] = None
    downloaded: bool = False

@dataclass
class CourseWorkMaterial:
    ID: str
    title: str
    creationTime: str
    dirname: str = None
    materials: List[Material] = field(default_factory=list)

@dataclass
class Course:
    ID: str
    title: str
    creationTime: str
    dirname: str = None
    courseWorkMaterials: List[CourseWorkMaterial] = field(default_factory=list)


def list_all_material(service_classroom) -> List[Course]:
    result = []
    courses = call_list_api(service_classroom.courses())
    for course in courses['courses']:
        course = Course(ID=course['id'], title=course['name'], creationTime=course['creationTime'])
        result.append(course)

        course_materials = call_list_api(service_classroom.courses().courseWorkMaterials(), courseId=course.ID)
        for course_material in course_materials['courseWorkMaterial']:
            cwm = CourseWorkMaterial(ID=course_material['id'], title=course_material['title'], creationTime=course_material['creationTime'])
            course.courseWorkMaterials.append(cwm)

            for material in course_material['materials']:
                if 'driveFile' not in material:
                    print(f"WARNING: unsupported material type {' '.join(material)}")
                    continue
                drivefile = material['driveFile']['driveFile']
                mat = Material(ID=drivefile['id'], title=drivefile['title'])
                cwm.materials.append(mat)

    return result


def title_to_filename(title):
    s = ''.join(
        ' ' if not c.isprintable()
        else '_' if c in '\\/<>|?*'
        else '-' if c == ':'
        else "'" if c == '"'
        else c
        for c in title)
    s = s.strip().strip('.')[:200]
    return s or "no_name"


def make_unique_names(names: List[str], has_extension):
    counter = Counter(names)
    suffix = {name:1 for name in counter if counter[name] > 1}
    if not suffix:
        return names
    else:
        new_names = []
        for name in names:
            if name in suffix:
                if has_extension:
                    # 'photo.jpg' --> 'photo_001.jpg'
                    s = name.split('.')
                    s1 = s[:-2]
                    s2 = s[-2:]
                    s2[0] += f"_{suffix[name]:03}"
                    n = '.'.join(s1 + s2)
                else:
                    # 'party' --> 'party_001'
                    n = f"{name}_{suffix[name]:03}"
                new_names.append(n)
                suffix[name] += 1
            else:
                new_names.append(name)
        return new_names


def assign_directory_names(collection):
    collection.sort(key=lambda item: item.creationTime)
    for item in collection:
        item.dirname = title_to_filename(item.title)
    new_names = make_unique_names([item.dirname for item in collection], has_extension=False)
    for item, new_name in zip(collection, new_names):
        item.dirname = new_name


def choose_mime_type(choices: List[str], document_type: str):
    fmt = [t for t in choices if 'vnd.openxmlformats-officedocument' in t] or \
        [t for t in choices if 'vnd.oasis.opendocument' in t] or \
        [t for t in choices if 'application/pdf' in t] or \
        [t for t in choices if t.startswith('image/')] or \
        choices
    if fmt:
        return fmt[0]
    if document_type == 'application/vnd.google-apps.document':
        return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    if document_type == 'application/vnd.google-apps.spreadsheet':
        return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    if document_type == 'application/vnd.google-apps.presentation':
        return 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    if document_type == 'application/vnd.google-apps.drawing':
        return 'application/pdf'
    else:
        return document_type


def assign_file_names(materials: List[Material], service_drive):
    # Download file metadata from Google Drive
    for material in materials:
        attr = service_drive.files().get(fileId=material.ID, fields='size,createdTime,mimeType,exportLinks').execute()
        material.size = attr.get('size')
        material.creationTime = attr['createdTime']
        material.mimeType = attr.get('mimeType')
        material.exportLinks = attr.get('exportLinks', {})
        material.filename = title_to_filename(material.title)
    # Sort files by creation time
    materials.sort(key=lambda item: item.creationTime)
    # Determine export format for Google documents and add extensions to filenames
    for material in materials:
        if not material.size:
            material.mimeType = choose_mime_type(list(material.exportLinks), material.mimeType)
        extensions = mimetypes.guess_all_extensions(material.mimeType or '')
        if extensions:
            if not any(ext for ext in extensions if material.filename.lower().endswith(ext.lower())):
                material.filename += mimetypes.guess_extension(material.mimeType)
    # Assign unique file names
    new_names = make_unique_names([material.filename for material in materials], has_extension=True)
    for material, new_name in zip(materials, new_names):
        material.filename = new_name


def assign_dir_and_file_names(courses: List[Course], basedir: Path, service_drive) -> int:
    count_to_refresh = 0
    count_to_download = 0
    # Sort directories and assign directory names
    assign_directory_names(courses)
    for course in courses:
        assign_directory_names(course.courseWorkMaterials)
        for course_material in course.courseWorkMaterials:
            path = basedir / course.dirname / course_material.dirname
            # If there is the right number of files in the directory, assume all the files
            # have already been downloaded
            if path.is_dir() and len(os.listdir(path)) >= len(course_material.materials):
                for material in course_material.materials:
                    material.downloaded = True
            else:
                count_to_refresh += 1
                for material in course_material.materials:
                    material.downloaded = False

    i = 0
    for course in courses:
        for course_material in course.courseWorkMaterials:
            if any(not material.downloaded for material in course_material.materials):
                print(f"\r{i}/{count_to_refresh}", end='', flush=True)
                assign_file_names(course_material.materials, service_drive)
                for material in course_material.materials:
                    material.downloaded = (basedir / course.dirname / course_material.dirname / material.filename).is_file()
                    if not material.downloaded:
                        count_to_download += 1
    print(f"\r{count_to_refresh}/{count_to_refresh}")
    return count_to_download


def download_file(material: Material, filepath: Path, service_drive, credentials):
    if material.size:
        # File can be downloaded as is from Drive
        data = service_drive.files().get_media(fileId=material.ID).execute()
    elif material.mimeType in material.exportLinks:
        # File can be downloaded from one of the availabile export links
        # https://stackoverflow.com/questions/40890534/google-drive-rest-api-files-export-limitation
        export_link = material.exportLinks[material.mimeType]
        r = requests.get(export_link, headers = {'Authorization': 'Bearer ' + credentials.token})
        if r.headers['Content-Type'] == material.mimeType:
            data = r.content
        else:
            print("WARNING: unable to download material")
    else:
        # File must be exported. Warning: there's a size limit on this conversion.
        data = service.files().export_media(fileId=material.ID, mimeType=material.mimeType).execute()

    filepath.parent.mkdir(parents=True, exist_ok=True)
    with open(filepath, 'xb') as f:
        f.write(data)


def download_missing_files(courses: List[Course], basedir: Path, total: int, service_drive, credentials):
    i = 1
    for course in courses:
        for course_material in course.courseWorkMaterials:
            for material in course_material.materials:
                if not material.downloaded:
                    relpath = Path(course.dirname) / course_material.dirname / material.filename
                    filepath = basedir / relpath
                    print(f"{i}/{total} {relpath}")
                    download_file(material, filepath, service_drive, credentials)
                    i += 1


def main():
    output_path = Path('../output')

    print("Authenticating...")
    creds = get_credentials()
    service_classroom = build('classroom', 'v1', credentials=creds)
    service_drive = build('drive', 'v3', credentials=creds)

    print("Reading courses...")
    courses = list_all_material(service_classroom)

    print("Retrieving file list...")
    total = assign_dir_and_file_names(courses, output_path, service_drive)

    print(f"Downloading...")
    download_missing_files(courses, output_path, total, service_drive, credentials=creds)


if __name__ == '__main__':
    main()
