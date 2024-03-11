## Imports
import pandas as pd
import os
import shutil
import json

## Global constants
org_name = "oandm"
images = "images"
index = "index.html"
lorem_ipsum = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."

## Function to convert "Demo Sprints" to "demo-sprints"
def convert(input_string):
    return "sprint-" + input_string.replace(' ', '-').lower()

## Read excel file
df = pd.read_excel('sprint-excel.xlsx', header=0)

## Extract rows
lab_titles = df.iloc[:,1]
sprint_titles = df.iloc[:,2]
descriptions = df.iloc[:,3]
estimated_time = df.iloc[:,4]
objective1 = df.iloc[:,5]
objective2 = df.iloc[:,6]
objective3 = df.iloc[:,7]
objective4 = df.iloc[:,8]
objective5 = df.iloc[:,9]
author = df.iloc[:,10]
contributors = df.iloc[:,11]
last_updated_by = df.iloc[:,12]
last_updated_date = df.iloc[:,13]

## .md file creation
def create_md_file(i) :

    title_A = f"# {sprint_titles[i]}"

    if not pd.isnull(estimated_time[i]):
        estimated_time_A = f"Estimated Time: {estimated_time[i]} minutes."
    else:
        estimated_time_A = f"Estimated Time: 10 minutes."

    objective1_A = objective1[i]
    objective2_A = objective2[i]
    objective3_A = objective3[i]
    objective4_A = objective4[i]
    objective5_A = objective5[i]

    content_title = f"{title_A}\n\n"
    content_estimated_time = f"{estimated_time_A}\n\n"
    content_objective1_A = f"## {objective1_A}\n\n"
    content_objective2_A = f"## {objective2_A}\n\n"
    content_objective3_A = f"## {objective3_A}\n\n"
    content_objective4_A = f"## {objective4_A}\n\n"
    content_objective5_A = f"## {objective5_A}\n\n"
    content_learn_more = f"## Learn More\n\n"
    content_acknowledgement = f"## Acknowledgements\n\n"
    content_last_updated_by_and_date_A = f"* **Author** - {author[i]}\n* **Contributors** - {contributors[i]}\n* **Last Updated By/Date** - {last_updated_by[i]}, {last_updated_date[i]}"
    content_learn_more_example = f"[link_text] (<link>)\n\n"

    content = content_title + content_estimated_time

    ## If any objective is null, do not print it
    if not pd.isnull(objective1_A):
        content += content_objective1_A
        content += f"{lorem_ipsum}\n\n"
    if not pd.isnull(objective2_A):
        content += content_objective2_A
        content += f"{lorem_ipsum}\n\n"
    if not pd.isnull(objective3_A):
        content += content_objective3_A
        content += f"{lorem_ipsum}\n\n"
    if not pd.isnull(objective4_A):
        content += content_objective4_A
        content += f"{lorem_ipsum}\n\n"
    if not pd.isnull(objective5_A):
        content += content_objective5_A
        content += f"{lorem_ipsum}\n\n"

    content += content_learn_more + content_learn_more_example + content_acknowledgement + content_last_updated_by_and_date_A

    return content

## manifest.json file creation
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
                "title": "Related Sprint: Title 1 ",
                "description": " ",
                "filename": " "
            },
            {
                "title": "Related Sprint: Title 2 ",
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

## Main function
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



