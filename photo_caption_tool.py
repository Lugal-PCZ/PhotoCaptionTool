import configparser
import csv
import math
import os
import re
from pathlib import Path
from datetime import datetime

import exif
from PIL import Image, ImageOps, ImageFont, ImageDraw
from docx import Document
from docx.shared import Mm


configs = configparser.ConfigParser()
images_directory = ""
all_images_exif_data = {}
rotation = [1, 8, 3, 6]  # Rotation of images, as represented in EXIF
valid_actions = []

# TODO: rename the internal-only function definitions with single underscores
# TODO: add paper size to the configs
# TODO: make CSV and Word doc viewing their own menu items
# TODO: force upper case for config variables


def facing(azimuth) -> str:
    if configs["FACING"]["precision"] == "coarse":
        increment = 22.5
        directions = [
            "N",
            "NE",
            "E",
            "SE",
            "S",
            "SW",
            "W",
            "NW",
            "N",
        ]
    elif configs["FACING"]["precision"] == "fine":
        increment = 11.25
        directions = [
            "N",
            "NNE",
            "NE",
            "ENE",
            "E",
            "ESE",
            "SE",
            "SSE",
            "S",
            "SSW",
            "SW",
            "WSW",
            "W",
            "WNW",
            "NW",
            "NNW",
            "N",
        ]
    try:
        if configs["FACING"]["precision"] == "precise":
            bearing = f"{round(azimuth)}°"
        else:
            bearing = directions[math.ceil(math.floor((azimuth % 360) / increment) / 2)]
    except:
        bearing = ""
    return bearing


def parse_gps_coords(lat, lon, ns, ew) -> str:
    try:
        latitude = f"{int(lat[0])}°{int(lat[1])}'{int(lat[2])}\"{ns}"
        longitude = f"{int(lon[0])}°{int(lon[1])}'{int(lon[2])}\"{ew}"
        return f"{latitude} {longitude}"
    except:
        return ""


def display_menu() -> str:
    global valid_actions
    if images_directory:
        print(f"\nPhotos Folder: {images_directory}")
    else:
        print(f"\nPhotos Folder: (not set)")
    print("----------------------------")
    print(" Choose an action:")
    print(" 1 - Load Photos from Folder")
    valid_actions = ["1"]
    if images_directory:
        print(" 2 - Create New Photo Log")
        valid_actions.append("2")
        csv_file = Path(images_directory) / "Photo Log.csv"
        if csv_file.is_file():
            print(" 3 - Update Original Photos")
            valid_actions.append("3")
            print(" 4 - Create Annotated Photos")
            valid_actions.append("4")
            print(" 5 - Create Contact Sheet")
            valid_actions.append("5")
    print(" Q - Quit")
    valid_actions.append("Q")
    print("----------------------------")
    return input().upper()


def view_csv_file(csv_file) -> None:
    if input("Type “Y” to open this file now. ").upper() == "Y":
        try:
            os.startfile(csv_file)
        except:
            os.system(f'open "{csv_file}"')


def load_photos() -> None:
    global images_directory
    global all_images_exif_data
    all_images_exif_data = {}
    images_directory = (
        input("Enter the Images Folder\n(type the path or drag the folder onto here): ")
        .strip("'")
        .strip('"')
        .strip()
    )
    if "PosixPath" in str(type(Path())):
        images_directory = images_directory.replace("\\", "")
    if images_directory.startswith("~"):
        images_directory = images_directory.replace("~", str(Path.home()), 1)
    images = []
    items = Path(images_directory).glob("*.*")
    for each_item in items:
        if re.match(".*\.jpe?g", each_item.name, flags=re.IGNORECASE):
            images.append(each_item.name)
            images.sort()
    if len(images) == 0:
        images_directory = ""
        print("NOTICE: There were no valid JPEG images in the selected directory.")
    else:
        for each_image in images:
            with open(Path(images_directory) / each_image, "rb") as f:
                the_image = exif.Image(f)
                all_images_exif_data[each_image] = the_image
                if not all_images_exif_data[each_image].has_exif:
                    all_images_exif_data[each_image].set("artist", "")
    main()


def create_csv() -> None:
    headers = [
        "Photo",
        "Photographer",
        "Project",
        "Site",
        "Timestamp",
        "GPS Coordinates",
        "Facing",
        "Description",
    ]
    if not all_images_exif_data:
        main()
    csv_file = Path(images_directory) / "Photo Log.csv"
    if (
        csv_file.is_file()
        and input(f"“{csv_file}” already exists. Type “Y” to overwrite it. ").upper()
        != "Y"
    ):
        view_csv_file(csv_file)
        main()
    # Get defaults
    data_for_csv = []
    for each_image in all_images_exif_data:
        image_data = {"Photo": each_image}
        image_data["Photographer"] = configs["DEFAULTS"]["photographer"]
        image_data["Project"] = configs["DEFAULTS"]["project"]
        image_data["Site"] = configs["DEFAULTS"]["site"]
        if not configs["DEFAULTS"]["photographer"]:
            try:
                image_data["Photographer"] = all_images_exif_data[each_image].get(
                    "artist"
                )
            except:
                pass
        try:
            image_data["Timestamp"] = str(
                datetime.strptime(
                    all_images_exif_data[each_image].get("datetime_original"),
                    "%Y:%m:%d %H:%M:%S",
                )
            )
        except:
            image_data["Timestamp"] = ""
        image_data["GPS Coordinates"] = parse_gps_coords(
            all_images_exif_data[each_image].get("gps_latitude"),
            all_images_exif_data[each_image].get("gps_longitude"),
            all_images_exif_data[each_image].get("gps_latitude_ref"),
            all_images_exif_data[each_image].get("gps_longitude_ref"),
        )
        image_data["Description"] = ""
        # Photo taken with iOS Camera.app:
        if all_images_exif_data[each_image].get("image_description"):
            image_data["Description"] = (
                all_images_exif_data[each_image].get("image_description").strip()
            )
        # Photo taken with Theodolite.app:
        if all_images_exif_data[each_image].get("user_comment"):
            image_data["Description"] = (
                all_images_exif_data[each_image].get("user_comment").strip()
            )
        image_data["Facing"] = facing(
            all_images_exif_data[each_image].get("gps_img_direction")
        )
        data_for_csv.append(image_data)
    with open(csv_file, "w", newline="") as f:
        csv_out = csv.DictWriter(f, fieldnames=headers)
        csv_out.writeheader()
        csv_out.writerows(data_for_csv)
    print("“Photo Log.csv” created.")
    view_csv_file(csv_file)
    main()


def update_originals() -> None:
    csv_file = Path(images_directory) / "Photo Log.csv"
    if not csv_file.is_file():
        main()
    with open(csv_file, "r") as f:
        reader = csv.DictReader(f)
        for each_photo in reader:
            # Update all_images_exif_data with the values from the CSV
            if each_photo["Photographer"]:
                all_images_exif_data[each_photo["Photo"]].set(
                    "artist", each_photo["Photographer"]
                )
            caption = ""
            if each_photo["Project"]:
                caption = f"Project: {each_photo['Project']}. "
            if each_photo["Site"]:
                caption += f"Site: {each_photo['Site']}. "
            caption += each_photo["Description"]
            all_images_exif_data[each_photo["Photo"]].set("image_description", caption)
            # Delete the Theodolite caption saved in the user_comment EXIF field
            if all_images_exif_data[each_photo["Photo"]].get("user_comment"):
                all_images_exif_data[each_photo["Photo"]].delete("user_comment")
            with open(Path(images_directory) / each_photo["Photo"], "wb") as f:
                f.write(all_images_exif_data[each_photo["Photo"]].get_file())
    main()


def annotate_photos() -> None:
    global all_images_exif_data
    csv_file = Path(images_directory) / "Photo Log.csv"
    if not csv_file.is_file():
        main()
    output_dir = Path(images_directory) / "Annotated Photos"
    if output_dir.is_dir():
        if (
            input(f"“{output_dir}” already exists. Type “Y” to overwrite it. ").upper()
            != "Y"
        ):
            main()
    output_dir.mkdir(exist_ok=True)
    with open(csv_file, "r") as f:
        reader = csv.DictReader(f)
        i = 1
        for each_photo in reader:
            print(f"{i}: Annotating photo {each_photo['Photo']}.")
            label = []
            line = [each_photo["Photo"]]
            if each_photo["Timestamp"]:
                line.append(each_photo["Timestamp"])
            label.append("   ".join(line))
            line = []
            if each_photo["Project"]:
                line.append(f"Project: {each_photo['Project']}")
            if each_photo["Site"]:
                line.append(f"Site: {each_photo['Site']}")
            if each_photo["Description"]:
                line.append(each_photo["Description"])
            if line:
                label.append(".  ".join(line))
            line = []
            if each_photo["Facing"]:
                line.append(f"Facing {each_photo['Facing']}")
            if each_photo["GPS Coordinates"]:
                line.append(each_photo["GPS Coordinates"])
            if line:
                label.append(".  ".join(line))
            if not label:
                label = ["(unlabeled)"]
            orientation = all_images_exif_data[each_photo["Photo"]].get("orientation")
            if not orientation:
                orientation = 1
            img = Image.open(Path(images_directory) / each_photo["Photo"])
            img = img.rotate(rotation.index(orientation) * 90, expand=True)
            img = ImageOps.pad(img, (img.width, img.height + 200), centering=(0, 0))
            img_annotation = ImageDraw.Draw(img, mode="RGB")
            try:
                thefont = ImageFont.truetype("Helvetica.ttc", 46)
            except:
                thefont = ImageFont.truetype("arial.ttf", 46)
            img_annotation.text(
                (20, img.height - 180),
                "\n".join(label),
                font=thefont,
                fill=(255, 255, 255),
                spacing=20,
            )
            filename = each_photo["Photo"].split(".")
            filename[-2] = f"{filename[-2]}_Annotated"
            filename = ".".join(filename)
            img.save(
                Path(output_dir) / filename,
                quality=80,
                optimize=True,
                progressive=True,
            )
            img.close()
            i += 1
    main()


def create_word_doc() -> None:
    csv_file = Path(images_directory) / "Photo Log.csv"
    if not csv_file.is_file():
        main()
    # Create a new Word document on A4 paper
    document = Document()
    section = document.sections[0]
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    section.left_margin = Mm(12)
    section.right_margin = Mm(12)
    section.top_margin = Mm(12)
    section.bottom_margin = Mm(12)
    csv_file = Path(images_directory) / "Photo Log.csv"
    with open(csv_file, "r") as f:
        reader = csv.DictReader(f)
        i = 1
        for each_photo in reader:
            print(f"{i}: Adding photo {each_photo['Photo']} to contact sheet.")
            document.add_picture(
                str(Path(images_directory) / each_photo["Photo"]), width=Mm(150)
            )
            label = [each_photo["Photo"]]
            if each_photo["Description"]:
                label.append(each_photo["Description"])
            if each_photo["Facing"]:
                label.append(f"Facing {each_photo['Facing']}")
            document.add_paragraph("\n".join(label))
            i += 1
    document.save(Path(images_directory) / "Contact Sheet.docx")
    print("“Contact Sheet.docx” created.")
    main()


def main() -> None:
    global configs
    try:
        with open("configs.ini", "r") as f:
            pass
    except:
        with open("configs.ini", "w") as f:
            f.write("[DEFAULTS]\n")
            f.write("photographer = \n")
            f.write("project = \n")
            f.write("site = \n\n")
            f.write("[FACING]\n")
            f.write("# options are:\n")
            f.write("#  'coarse' (N, NE, E, SE, S, SW, W, NW)\n")
            f.write("#  'fine' (N, NNE, NE, ENE, E, and so-on)\n")
            f.write("#  'precise' (the actual bearing, in degrees)\n")
            f.write("precision = coarse\n")
    configs.read("configs.ini")
    action = ""
    while action not in valid_actions:
        action = display_menu()
    if action == "1":
        load_photos()
    elif action == "2":
        create_csv()
    elif action == "3":
        update_originals()
    elif action == "4":
        annotate_photos()
    elif action == "5":
        create_word_doc()
    elif action == "Q":
        exit()


if __name__ == "__main__":
    main()
