## Imports
import pandas as pd
import os
import shutil
import json
import random

## Global constants
org_name = "oandm"
images = "images"
index = "index.html"
lorem_ipsum = "<!-- Replace below text with instructions -->\nLorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."

## Read excel file
df = pd.read_excel('sprint-excel.xlsx', header=0)

max_columns = 100

## Extract rows
generate_sprint = df.iloc[:,1]
lab_titles = df.iloc[:,2]
sprint_titles = df.iloc[:,3]
descriptions = df.iloc[:,4]
estimated_time = df.iloc[:,5]
author = df.iloc[:,6]
contributors = df.iloc[:,7]
last_updated_by = df.iloc[:,8]
last_updated_date = df.iloc[:,9]

## Function to convert "Demo Sprints" to "sprint-demo-sprints"
def sprint_title_converter(input_string: str):
    return "sprint-" + input_string.replace(' ', '-').lower()

## Function to convert "Using Compass Icon" to "using-compass-icon"
def image_name_converter(input_string: str):
    return input_string.replace(' ', '-').lower()

## Generate image examples
def generate_image_examples(count: int, step_name: str):
    image_examples = ""
    for i in range(count):
        image_name = f"{image_name_converter(step_name)}_{i+1}"
        image_example = f"<!-- Rename the image as {image_name} and add it in images folder -->\n"
        image_example += f"![{step_name} {i+1}](./images/{image_name}.png)\n\n"
        image_examples += image_example
    return image_examples

## Function to convert code into code snippet
def code_snippet_convert(code):
    code_snippet = f"```text\n<copy>\n    {code}\n</copy>\n```\n\n"
    return code_snippet

## Function to create content for each objective
def create_content_for_each_objective(objective: str, instructions, images_count: int, code_snippet):
    if pd.isnull(objective):
        objective = f"Objective_{random.randint(1,100)}"
    content = f"## {objective}\n\n"
    if not pd.isnull(instructions):
        content += f"{instructions}\n\n"
    else:
        content += f"{lorem_ipsum}\n\n"
    if not pd.isnull(images_count):
        content += generate_image_examples(images_count, objective)
    if not pd.isnull(code_snippet):
        content += code_snippet_convert(code_snippet)
    return content

## Function to create a video link example
def create_video_link():
    content = ""
    content += "<!-- (Embed Video)\nEmbedding a video from Oracle Video Hub (recommended):\n"
    content += "    For example, if the link of Oracle Video Hub is https://videohub.oracle.com/media/Getting+Started+with+SuiteCommerce+Themes/1_rd0x9j2i, then code snippet to embed the video will be,\n"
    content += "[Alternate text for Video](videohub:1_rd0x9j2i)\n"
    content += "Embedding a video from YouTube:\n"
    content += "    For example, if the link of YouTube video is https://www.youtube.com/watch?v=dQw4w9WgXcQv, then code snippet to embed the video will be,\n"
    content += "[Alternate text for Video](youtube:dQw4w9WgXcQ)\n"
    content += "-->\n\n"
    return content

## Function to fill main content
def main_content(row):
    step_name = 'step_name'
    step_instructions = "step_instructions"
    step_images_count = "step_images_count"
    step_codesnippet = "step_codesnippet"
    content = ""
    for i in range(int(max_columns/4)):
        if pd.isnull(step_name) and pd.isnull(step_instructions) and pd.isnull(step_images_count) and pd.isnull(step_codesnippet):
            break
        step_name = df.iloc[:,10+i*4][row]
        step_instructions = df.iloc[:,11+i*4][row]
        step_images_count = df.iloc[:,12+i*4][row]
        if not pd.isnull(step_images_count):
            step_images_count = int(step_images_count)
        step_codesnippet = df.iloc[:,13+i*4][row]
        if pd.isnull(step_name) and pd.isnull(step_instructions) and pd.isnull(step_images_count) and pd.isnull(step_codesnippet):
            break
        content += create_content_for_each_objective(step_name, step_instructions, step_images_count, step_codesnippet)
    return content

## .md file creation
def create_md_file(i) :

    title_A = f"# {sprint_titles[i]}"

    if not pd.isnull(estimated_time[i]):
        estimated_time_A = f"Estimated Time: {int(estimated_time[i])} minutes."
    else:
        estimated_time_A = f"Estimated Time: 10 minutes."

    content_title = f"{title_A}\n\n"
    content_estimated_time = f"{estimated_time_A}\n\n"
    content_learn_more = f"## Learn More\n\n"
    content_acknowledgement = f"## Acknowledgements\n\n"
    content_last_updated_by_and_date_A = f"* **Author** - {author[i]}\n* **Contributors** - {contributors[i]}\n* **Last Updated By/Date** - {last_updated_by[i]}, {last_updated_date[i]}"

    content_learn_more_example = "<!-- (Add links)\n\n"
    content_learn_more_example += "[link_text] (<link>)\n\n"
    content_learn_more_example += "[link_text] (<link>)\n\n"
    content_learn_more_example += "Replace link_text and link, for example:\n[Logging Analytics](<https://docs.oracle.com/en-us/iaas/logging-analytics/home.htm>)\n"
    content_learn_more_example += "-->\n\n"

    content = content_title + content_estimated_time + create_video_link()

    content += main_content(i)

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

## Main function
def main_function():
    if not os.path.exists(org_name):
        os.makedirs(org_name)

    for i in range(len(df)):

        if generate_sprint[i] == "Yes":
            lab_title = sprint_title_converter(lab_titles[i])
            folder_path_sprint = os.path.join(org_name, lab_title)

            if not os.path.exists(folder_path_sprint) :
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

main_function()









