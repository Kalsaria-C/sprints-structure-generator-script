import pandas as pd
import os
import shutil
import json

org_name = "oandm"
images = "images"
index = "index.html"

def convert(input_string):
    return "sprint-" + input_string.replace(' ', '-').lower()

df = pd.read_excel('sprint-excel.xlsx', header=0)

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

if not os.path.exists(org_name):
    os.makedirs(org_name)

for i in range(len(lab_titles)):

    lab_title = convert(lab_titles[i])

    folder_path_sprint = os.path.join(org_name, lab_title)
    os.makedirs(folder_path_sprint)
    folder_path_images = os.path.join(folder_path_sprint, images)
    os.makedirs(folder_path_images)
    md_file_name = f"{lab_title}.md"
    md_file_path = os.path.join(folder_path_sprint, md_file_name)
    with open(md_file_path, "w") as file:
        file.write(create_md_file(i))

    index_file_path = os.path.join(folder_path_sprint, index)
    shutil.copyfile('./index.html', index_file_path)
    manifest = "manifest.json"
    manifest_file_path = os.path.join(folder_path_sprint, manifest)

    manifest_file_name = f"../{lab_title}/{lab_title}.md"
    with open(manifest_file_path, "w") as file:
        file.write(create_manifest_file(sprint_titles[i], descriptions[i], manifest_file_name))



