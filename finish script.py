# -*- coding: utf-8 -*-
"""
Created on Mon Jun 16 11:35:58 2025

@author: Viera Rattayová, rattayovaviera@gmail.com, https://github.com/VierkaOdjazierka/
"""

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os

import textwrap


#################################################################################
#################################################################################
# set the folders and files

# === CONFIGURATION ===



csv_path = "C://Users//Viera//Desktop//menovky//menovky_automaticke//na menovky.csv"
background_path = "C:/Users/Viera/Desktop/menovky/menovky_automaticke/menovka.png"  # <-- 1063x591px at 300 DPI
output_folder = "C:/Users/Viera/Desktop/menovky/output_tags/"

# set dpi
dpi = (300, 300)

##################################
#set font used in tags

font_path_bold = "C:\\Windows\\Fonts\\arialbd.ttf"     # Arial Bold
font_path_regular = "C:\\Windows\\Fonts\\arial.ttf"    # Arial Regular


##################################
#size of text in name tags, separate for name, institucion....
font_size_name = 40
font_size_institution = 20
font_size_country = 20

#font thiskness
font_thick_name = font_path_bold
font_thick_institution = font_path_regular
font_thick_country = font_path_bold
##################################

n_char=40   # set number of characters per line, after this number the text will be wrapet to other line
            # this is important to change, if you change font size to bigger
            # to fit it to size of name tag
##################################
# set number of tags without text, when someone without registration come to conference....

number_of_free_tags=5
##################################
# set position of text in tags, separate neam, institution....
name_position=110
inst_position=180
count_position=240


##################################
# printing size of name tags
tag_width_cm = 9.4
tag_height_cm = 5.7


#################################################################################
#################################################################################
# function for wraping long text in institucion name

def center_wrapped_text(draw, text, font, y_start, line_spacing=10, max_width=950):
    """
    Draw wrapped text centered on the image.
    """
    # Wrap text to fit within max_width
    lines = []
    for line in text.split('\n'):
        wrapped = textwrap.wrap(line, width=n_char)  # set number of characters per line
        lines.extend(wrapped if wrapped else [" "])

    # Draw each line centered
    for i, line in enumerate(lines):
        bbox = draw.textbbox((0, 0), line, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        x = (W - w) // 2
        y = y_start + i * (h + line_spacing)
        draw.text((x, y), line, font=font, fill="black")



#################################################################################
#################################################################################


# === CREATE OUTPUT FOLDER ===
os.makedirs(output_folder, exist_ok=True)

# === LOAD DATA ===
df = pd.read_csv(csv_path)
df.columns = ["Name", "Institution", "Country"] #rename columns in file as script need


full_nan_rows = df[df.isnull().all(axis=1)]
keep_nan_rows = full_nan_rows.head(number_of_free_tags)
non_nan_rows = df[~df.isnull().all(axis=1)]
df = pd.concat([non_nan_rows, keep_nan_rows], ignore_index=True)
df = df.fillna("")



# === GENERATE TAGS ===
for idx, row in df.iterrows():
    name = row["Name"]
    institution = row["Institution"]
    country = row["Country"]

    # Load background
    bg = Image.open(background_path).convert("RGBA")
    draw = ImageDraw.Draw(bg)

    # Load fonts
    font_name = ImageFont.truetype(font_thick_name, font_size_name)
    font_inst = ImageFont.truetype(font_thick_institution, font_size_institution)
    font_country = ImageFont.truetype(font_thick_country, font_size_country)    # set font

    W, H = bg.size

    def center_text(text, font, y):
        w, h = draw.textsize(text, font=font)
        x = (W - w) // 2
        draw.text((x, y), text, font=font, fill="black")

    # Draw text (adjust Y positions as per your layout)
    center_wrapped_text(draw, name, font_name, y_start=name_position)
    center_wrapped_text(draw, institution, font_inst, y_start=inst_position)
    center_wrapped_text(draw, country, font_country, y_start=count_position)


    # Save with correct DPI
    out_path = os.path.join(output_folder, f"{idx+1:02d}_{name.replace(' ', '_')}.png")
    bg.save(out_path, dpi=dpi)

print(f"✅ {len(df)} tags generated in folder: {output_folder}")


#################################################################################
#################################################################################
#Final PDF

from PIL import Image
import os
import math

dpi = 300

# === CONFIGURATION ===
input_folder = "C:/Users/Viera/Desktop/menovky/output_tags"
output_pdf_path = "C:/Users/Viera/Desktop/menovky/name_tags_A4.pdf"



# Convert cm to pixels
tag_w = int(tag_width_cm / 2.54 * dpi)  # ~1063
tag_h = int(tag_height_cm / 2.54 * dpi)  # ~591

# A4 dimensions at 300 DPI
a4_w = int(21 / 2.54 * dpi)  # 2480
a4_h = int(29.7 / 2.54 * dpi)  # 3508

# Number of tags per row and column
tags_per_row = 2
tags_per_col = 4
tags_per_page = tags_per_row * tags_per_col

# Get tag images
tag_files = [f for f in sorted(os.listdir(input_folder)) if f.lower().endswith(".png")]
pages = []

# Loop through chunks of 8 images (one A4 page per chunk)
for i in range(0, len(tag_files), tags_per_page):
    chunk = tag_files[i:i+tags_per_page]
    page = Image.new("RGB", (a4_w, a4_h), "white")
    draw = ImageDraw.Draw(page)  # Create draw object once per page

    for idx, tag_file in enumerate(chunk):
        tag_img = Image.open(os.path.join(input_folder, tag_file)).convert("RGB")
        tag_img = tag_img.resize((tag_w, tag_h))

        col = idx % tags_per_row
        row = idx // tags_per_row

        margin_x = (a4_w - (tags_per_row * tag_w)) // (tags_per_row + 1)
        margin_y = (a4_h - (tags_per_col * tag_h)) // (tags_per_col + 1)

        x = margin_x + col * (tag_w + margin_x)
        y = margin_y + row * (tag_h + margin_y)

        # Paste the tag
        page.paste(tag_img, (x, y))

        # === Draw border ===
        border_thickness = 4
        draw.rectangle([ (x, y), (x + tag_w, y + tag_h) ], outline="black", width=border_thickness)

        # === Crop marks ===
        crop_len = 20
        crop_color = "black"
        draw.line([(x, y), (x + crop_len, y)], fill=crop_color, width=1)
        draw.line([(x, y), (x, y + crop_len)], fill=crop_color, width=1)
        draw.line([(x + tag_w, y), (x + tag_w - crop_len, y)], fill=crop_color, width=1)
        draw.line([(x + tag_w, y), (x + tag_w, y + crop_len)], fill=crop_color, width=1)
        draw.line([(x, y + tag_h), (x + crop_len, y + tag_h)], fill=crop_color, width=1)
        draw.line([(x, y + tag_h), (x, y + tag_h - crop_len)], fill=crop_color, width=1)
        draw.line([(x + tag_w, y + tag_h), (x + tag_w - crop_len, y + tag_h)], fill=crop_color, width=1)
        draw.line([(x + tag_w, y + tag_h), (x + tag_w, y + tag_h - crop_len)], fill=crop_color, width=1)

    # ✅ Add page to list after placing all tags
    pages.append(page)




# Save all pages as a PDF
pages[0].save(output_pdf_path, save_all=True, append_images=pages[1:])
print(f"✅ PDF created: {output_pdf_path}")