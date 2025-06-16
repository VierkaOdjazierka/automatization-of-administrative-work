# -*- coding: utf-8 -*-
"""
Created on Mon Jun 16 11:35:58 2025

@author: Viera Rattayová, rattayovaviera@gmail.com , https://github.com/VierkaOdjazierka/
"""

import os
import textwrap
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

#################################################################################
#################################################################################
#input paths

csv_path = "C://Users//Viera//Desktop//menovky//na menovky.csv"
output_folder = "C:/Users/Viera/Desktop/menovky/output_tag_table/"
font_path_bold = "C:\\Windows\\Fonts\\arialbd.ttf"
font_path_regular = "C:\\Windows\\Fonts\\arial.ttf"

#################################################################################
#################################################################################
# Font sizes
font_size_name = 120
font_size_institution = 80
font_size_country = 80

# Font weights
font_thick_name = font_path_bold
font_thick_institution = font_path_regular
font_thick_country = font_path_bold


#################################################################################
#################################################################################
# Max characters per line
n_char = 40

#################################################################################
#################################################################################
# Number of blank tags
number_of_free_tags = 5

#################################################################################
#################################################################################
# Text positions
name_position = 1500
inst_position = 1650
count_position = 1950

#################################################################################
#################################################################################

# Tag size in cm (A5 format)
tag_width_cm = 14.8  # A5 width
tag_height_cm =  21.0  # A5 height

# DPI settings
dpi_value = 300
dpi = (dpi_value, dpi_value)


#################################################################################
#################################################################################

# === HELPER FUNCTION ===
def center_wrapped_text(draw, text, font, y_start, W, line_spacing=10):
    lines = []
    for line in text.split('\n'):
        wrapped = textwrap.wrap(line, width=n_char)
        lines.extend(wrapped if wrapped else [" "])
    for i, line in enumerate(lines):
        bbox = draw.textbbox((0, 0), line, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        x = (W - w) // 2
        y = y_start + i * (h + line_spacing)
        draw.text((x, y), line, font=font, fill="black")


#################################################################################
#################################################################################

os.makedirs(output_folder, exist_ok=True)

#################################################################################
#################################################################################


# === LOAD DATA ===
df = pd.read_csv(csv_path)
df.columns = ["Name", "Institution", "Country"]
full_nan_rows = df[df.isnull().all(axis=1)]
keep_nan_rows = full_nan_rows.head(number_of_free_tags)
non_nan_rows = df[~df.isnull().all(axis=1)]
df = pd.concat([non_nan_rows, keep_nan_rows], ignore_index=True).fillna("")

# Calculate blank image size in pixels for A5
tag_w = int(tag_width_cm / 2.54 * dpi_value)
tag_h = int(tag_height_cm / 2.54 * dpi_value)

# === GENERATE TAGS ===
for idx, row in df.iterrows():
    name, institution, country = row["Name"], row["Institution"], row["Country"]
    bg = Image.new("RGB", (tag_w, tag_h), "white")
    draw = ImageDraw.Draw(bg)
    font_name = ImageFont.truetype(font_thick_name, font_size_name)
    font_inst = ImageFont.truetype(font_thick_institution, font_size_institution)
    font_country = ImageFont.truetype(font_thick_country, font_size_country)
    W, H = bg.size
    center_wrapped_text(draw, name, font_name, name_position, W)
    center_wrapped_text(draw, institution, font_inst, inst_position, W)
    center_wrapped_text(draw, country, font_country, count_position, W)
    safe_name = ''.join(c if c.isalnum() else '_' for c in name.strip())
    out_path = os.path.join(output_folder, f"{idx+1:02d}_{safe_name}.png")
    bg.save(out_path, dpi=dpi)

print(f"✅ {len(df)} tags generated in folder: {output_folder}")


##########################################################################


# === CREATE A4 LANDSCAPE PDF with 2 A5 TAGS SIDE BY SIDE ===
from PIL import Image

output_pdf_path = "C:/Users/Viera/Desktop/menovky/name_tags_A4_landscape.pdf"
a4_w = int(29.7 / 2.54 * dpi_value)  # A4 width in landscape
a4_h = int(21.0 / 2.54 * dpi_value)  # A4 height in landscape

tag_files = [f for f in sorted(os.listdir(output_folder)) if f.lower().endswith(".png")]
pages = []

for i in range(0, len(tag_files), 2):
    chunk = tag_files[i:i + 2]
    page = Image.new("RGB", (a4_w, a4_h), "white")
    draw = ImageDraw.Draw(page)

    for idx, tag_file in enumerate(chunk):
        tag_img = Image.open(os.path.join(output_folder, tag_file)).convert("RGB")
        tag_img = tag_img.resize((tag_w, tag_h))

        x = idx * (a4_w // 2) + ((a4_w // 2 - tag_w) // 2)
        y = (a4_h - tag_h) // 2

        page.paste(tag_img, (x, y))

        # Border
        border_thickness = 4
        draw.rectangle([ (x, y), (x + tag_w, y + tag_h) ], outline="black", width=border_thickness)

    # Vertical cutting line in the middle
    middle_x = a4_w // 2
    draw.line([(middle_x, 0), (middle_x, a4_h)], fill="black", width=1)

    pages.append(page)

if pages:
    pages[0].save(output_pdf_path, save_all=True, append_images=pages[1:])
    print(f"✅ A4 landscape PDF created with 2 tags per page: {output_pdf_path}")
else:
    print("⚠️ No tags found to include in A4 PDF.")