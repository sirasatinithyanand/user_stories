# -*- coding: utf-8 -*-
import pandas as pd
import os
from openai import OpenAI
import random
from tqdm import tqdm

# Initialize OpenAI client with your API key
api_key = "your gpt key"
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY_GPT_4") or api_key)

# Function to generate 4 to 6 user stories based on the feature description
def generate_user_stories(description, temperature, global_story_counter):
    user_stories = []

    messages = [
        {"role": "system", "content": (
            "You are a helpful assistant who generates agile user stories. "
            "Generate between 4 to 6 user stories based on the feature description and make sure the user stories are as diverse as possible. "
            "Do not generate User Story 1:, User Story 2:, etc. numbering in the output. "
            "Please restrict the number of user stories per cell to 1."
        )},
        {"role": "user", "content": f"Feature Description: {description}"}
    ]

    completion = client.chat.completions.create(
        model="gpt-4",
        messages=messages,
        temperature=temperature
    )

    stories = completion.choices[0].message.content.split('\n\n')
    for story in stories:
        if story.strip():
            # Include the same global_story_counter in the User Story and the User Stories ID
            user_stories.append((f"User Story {global_story_counter}:\n{story.strip()}", global_story_counter))
            global_story_counter += 1

    return user_stories, global_story_counter

# Add user stories to the copied DataFrame
def add_user_stories_to_df(df, global_story_counter):
    user_stories_list = []
    for _, row in tqdm(df.iterrows(), total=df.shape[0], desc="Generating User Stories"):
        feature_id = row['Feature ID']
        feature_description = row['Feature Description']
        stories, global_story_counter = generate_user_stories(feature_description, temperature=0.7, global_story_counter=global_story_counter)
        for story, counter in stories:
            user_story_id = f"{feature_id}-US-{counter}"
            user_stories_list.append({
                **row,
                'User Stories ID': user_story_id,
                'User Stories': story
            })
    return pd.DataFrame(user_stories_list), global_story_counter

# Process all Excel files in the given folder
input_folder_path = 'input_file path'
output_folder_path = 'output folder path'

# Create the output folder if it doesn't exist
os.makedirs(output_folder_path, exist_ok=True)

for filename in os.listdir(input_folder_path):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(input_folder_path, filename)

        # Reset the global story counter for each file
        global_story_counter = 1

        # Load Excel file and detect sheet names
        xls = pd.ExcelFile(file_path)
        sheet_name = xls.sheet_names[0]  # Dynamically select the first sheet

        # Load the DataFrame
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Generate user stories
        df_with_stories, global_story_counter = add_user_stories_to_df(df, global_story_counter)

        # Save the output to the new folder with the same name as the input file
        output_path = os.path.join(output_folder_path, filename)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            df_with_stories.to_excel(writer, sheet_name='User Stories', index=False)

