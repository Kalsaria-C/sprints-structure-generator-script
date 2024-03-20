import os, shutil, json, argparse, pandas as pd, glob
from urllib import request, error as urllibError

# CLI Arguments
argParser = argparse.ArgumentParser()
argParser.add_argument("-f", "--file", help="Path of the xlsx file to upload", required=True)
args = argParser.parse_args()

org_name = "oandm"
images = "images"
index = "index.html"
index_file_remote_url = 'https://raw.githubusercontent.com/oracle-livelabs/sprints/master/sample-sprints/sprint/index.html'
use_local_index_file = False    # Set this flag to true if you do not want to fetch the latest version of index.html from the oracle official sprints repo

def convert(input_string):
    return "sprint-" + input_string.replace(' ', '-').lower()

df = pd.read_excel(args.file, header=0)

lab_titles = df.iloc[:,1]
sprint_titles = df.iloc[:,2]
descriptions = df.iloc[:,3]
estimated_time = df.iloc[:,4]
objective1 = df.iloc[:,5]
objective2 = df.iloc[:,6]
objective3 = df.iloc[:,7]
objective4 = df.iloc[:,8]
objective5 = df.iloc[:,9]

def create_md_file(i) :
    title_A = f"# {sprint_titles[i]}"
    estimated_time_A = f"Estimated Time: {estimated_time[i]} minutes."
    objective1_A = f"## {objective1[i]}"
    objective2_A = f"## {objective2[i]}"
    objective3_A = f"## {objective3[i]}"
    objective4_A = f"## {objective4[i]}"
    objective5_A = f"## {objective5[i]}"
    learn_more = "## Learn More"
    acknowledgement = "## Acknowledgements"


    content = f"{title_A}\n\n{estimated_time_A}\n\n{objective1_A}\n\n{objective2_A}\n\n{objective3_A}\n\n{objective4_A}\n\n{objective5_A}\n\n{learn_more}\n\n{acknowledgement}"
    return content

index = "index.html"

def create_manifest_file(sprint_title, description, filename) :

    data = '''
    {
    "workshoptitle": "LiveLabs Sprints",
    "help": "livelabs-help-sprints_us@oracle.com",
    "tutorials": [
            {
                "title": "How to create alerts on logs with Logging Analytics?",
                "description": "Explains how to create OCI Alarms and verify it.",
                "filename": "sprint-alerts-on-logs-with-logging-analytics.md"
            },
            {
                "title": "Related Sprint: Title1 ",
                "description": " ",
                "filename": " "
            },
            {
                "title": "Related Sprint: Title2 ",
                "description": " ",
                "filename": " "
            }
        ],
        "task_type": "Sections"
    }
    '''

    data = json.loads(data)
    data["tutorials"][0]["title"] = sprint_title
    description = data["tutorials"][0]["description"] = description
    data["tutorials"][0]["filename"] = filename
    return json.dumps(data)


def get_latest_index_file_from_github(url:str='https://raw.githubusercontent.com/oracle-livelabs/sprints/master/sample-sprints/sprint/index.html'):
    """Checks that an index file is available in the remote repository and downloads it or shows a warning and uses the one actually available"""

    try:
        request.urlretrieve(url, 'latest/index.html')   # Download the latest index file from the remote repo (url param), no need to check or delete the cached index.html because it overrides it automatically
        print('\033[32m\033[1m[Info]: fetched latest version of index.html file\033[0m')
        return True

    except urllibError.HTTPError as e:
        print('\033[33m\033[1m[Warning]:', f'''
{e.code}: "{e.msg}" Error has occurred when fetching latest version of "index.html file" from remote github repo "{url}"
The script will proceed with local version of "index.html" file!
\033[0m''')
        return False


src_index_file = 'index.html'               # Default local index.html to copy from in case the remote repo is down or can not be fetched for any reason
if((not use_local_index_file) and get_latest_index_file_from_github(index_file_remote_url)):    # Fetch the latest index.html from the remote github repo
    src_index_file = 'latest/index.html'


if(glob.glob("./oandm")):
    # i = str(len(glob.glob("./oandm-save-*"))+1)
    i = max([
            int(x.split('-')[-1])
            for x in glob.glob("./oandm-save-*") 
            if x.split('-')[-1].isnumeric()
        ]) + 1
    os.rename("./oandm", "./oandm-save-" + str(i))


# Create base oandm dist directory
if not os.path.exists(org_name):
    os.makedirs(org_name)


for i in range(len(lab_titles)):

    lab_title = convert(lab_titles[i])

    # Create project folder
    folder_path_sprint = os.path.join(org_name, lab_title)
    os.makedirs(folder_path_sprint)

    # Create images folder
    folder_path_images = os.path.join(folder_path_sprint, images)
    os.makedirs(folder_path_images)

    # Create md file
    md_file_name = f"{lab_title}.md"
    md_file_path = os.path.join(folder_path_sprint, md_file_name)
    with open(md_file_path, "w") as file:
        file.write(create_md_file(i))

    # Copy index file to project
    index_file_path = os.path.join(folder_path_sprint, index)
    shutil.copyfile(src_index_file, index_file_path)

    # Create manifest file to project
    manifest = "manifest.json"
    manifest_file_path = os.path.join(folder_path_sprint, manifest)

    manifest_file_name = f"../{lab_title}/{lab_title}.md"
    with open(manifest_file_path, "w") as file:
        file.write(create_manifest_file(sprint_titles[i], descriptions[i], manifest_file_name))

    print(f'[Info]: Lab {lab_title} was created successfully!')
