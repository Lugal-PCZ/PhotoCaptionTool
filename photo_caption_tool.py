import configparser
import csv
import math
import os
import re
import subprocess
from pathlib import Path

from PIL import Image, ImageOps, ImageFont, ImageDraw
from docx import Document
from docx.shared import Inches, Mm


configs = configparser.ConfigParser()
images_directory = ""
all_images_exif_data = {}
rotation = ["1", "8", "3", "6"]  # Rotation of images, as represented in EXIF
valid_actions = []


def _display_menu() -> str:
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
        csv_file = Path(images_directory) / "Photo Log.csv"
        word_doc = Path(images_directory) / "Contact Sheet.docx"
        print(" 2 - Create New Photo Log")
        valid_actions.append("2")
        if csv_file.is_file():
            print(" 3 - View CSV File")
            valid_actions.append("3")
            print(" 4 - Create Annotated Photos")
            valid_actions.append("4")
            print(" 5 - Create Contact Sheet")
            valid_actions.append("5")
            if word_doc.is_file():
                print(" 6 - View Contact Sheet")
                valid_actions.append("6")
            print(" 7 - Update Original Photos")
            valid_actions.append("7")
    print(" Q - Quit")
    valid_actions.append("Q")
    print("----------------------------")
    return input().upper()


def _facing(azimuth) -> str:
    if configs["FACING"]["precision"].lower() == "coarse":
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
    elif configs["FACING"]["precision"].lower() == "fine":
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
        azimuth = float(azimuth)
        if configs["FACING"]["precision"].lower() == "precise":
            bearing = f"{round(azimuth)}°"
        else:
            bearing = directions[math.ceil(math.floor((azimuth % 360) / increment) / 2)]
    except:
        bearing = ""
    return bearing


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
            orientation = all_images_exif_data[each_photo["Photo"]]["orientation"]
            if not orientation:
                orientation = "1"
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
        main()
    data_for_csv = []
    for each_image in all_images_exif_data:
        image_data = {"Photo": each_image}
        image_data["Photographer"] = configs["DEFAULTS"]["photographer"]
        image_data["Project"] = configs["DEFAULTS"]["project"]
        image_data["Site"] = configs["DEFAULTS"]["site"]
        if not configs["DEFAULTS"]["photographer"]:
            photographer = []
            if all_images_exif_data[each_image]["artist"]:
                photographer.append(all_images_exif_data[each_image]["artist"])
            if (
                all_images_exif_data[each_image]["creator"]
                and all_images_exif_data[each_image]["creator"] != photographer[0]
            ):
                photographer.append(all_images_exif_data[each_image]["creator"])
            image_data["Photographer"] = ", ".join(photographer)
        image_data["Timestamp"] = all_images_exif_data[each_image]["datetimeoriginal"]
        image_data["GPS Coordinates"] = all_images_exif_data[each_image]["gpsposition"]
        image_data["Description"] = ""
        # Photo taken with iOS Camera.app:
        if all_images_exif_data[each_image]["imagedescription"]:
            image_data["Description"] = all_images_exif_data[each_image][
                "imagedescription"
            ].strip()
        # Photo taken with Theodolite.app:
        if all_images_exif_data[each_image]["usercomment"]:
            image_data["Description"] = all_images_exif_data[each_image][
                "usercomment"
            ].strip()
        image_data["Facing"] = _facing(
            all_images_exif_data[each_image]["gpsimgdirection"]
        )
        data_for_csv.append(image_data)
    with open(csv_file, "w", newline="") as f:
        csv_out = csv.DictWriter(f, fieldnames=headers)
        csv_out.writeheader()
        csv_out.writerows(data_for_csv)
    print("“Photo Log.csv” created.")
    main()


def create_word_doc() -> None:
    csv_file = Path(images_directory) / "Photo Log.csv"
    if not csv_file.is_file():
        main()
    word_doc = Path(images_directory) / "Contact Sheet.docx"
    if (
        word_doc.is_file()
        and input(f"“{word_doc}” already exists. Type “Y” to overwrite it. ").upper()
        != "Y"
    ):
        main()
    # Create a new Word document on A4 paper
    document = Document()
    section = document.sections[0]
    if configs["DEFAULTS"]["papersize"].lower() == "a4":
        section.page_height = Mm(297)
        section.page_width = Mm(210)
        section.left_margin = Mm(12)
        section.right_margin = Mm(12)
        section.top_margin = Mm(12)
        section.bottom_margin = Mm(12)
        photowidth = Mm(100)
    elif configs["DEFAULTS"]["papersize"].lower() == "letter":
        section.page_height = Inches(11)
        section.page_width = Inches(8.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        photowidth = Inches(4)
    with open(csv_file, "r") as f:
        reader = csv.DictReader(f)
        i = 1
        for each_photo in reader:
            print(f"{i}: Adding photo {each_photo['Photo']} to contact sheet.")
            document.add_picture(
                str(Path(images_directory) / each_photo["Photo"]), width=photowidth
            )
            label = [each_photo["Photo"]]
            if each_photo["Description"]:
                label.append(each_photo["Description"])
            if each_photo["Facing"]:
                label.append(f"Facing {each_photo['Facing']}")
            document.add_paragraph("\n".join(label))
            i += 1
    document.save(word_doc)
    print("“Contact Sheet.docx” created.")
    main()


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
        tags = [
            "-datetimeoriginal",
            "-artist",
            "-creator",
            "-imagedescription",
            "-usercomment",
            "-gpsposition",
            "-gpsimgdirection",
            "-orientation#",
        ]
        for each_image in images:
            exif_data = (
                subprocess.run(
                    [
                        Path(configs["DEFAULTS"]["exiftool"]),
                        "-T",
                        "-c",
                        "%d°%d'%.2f\"",
                        *tags,
                        Path(images_directory) / each_image,
                    ],
                    capture_output=True,
                    text=True,
                )
                .stdout.strip()
                .split("\t")
            )
            all_images_exif_data[each_image] = dict(
                zip(
                    [x.strip("-").strip("#") for x in tags],
                    ["" if x == "-" else x for x in exif_data],
                )
            )
    main()


def update_originals() -> None:
    csv_file = Path(images_directory) / "Photo Log.csv"
    if not csv_file.is_file():
        main()
    with open(csv_file, "r") as f:
        reader = csv.DictReader(f)
        for each_photo in reader:
            caption = ""
            if each_photo["Project"]:
                caption = f"Project: {each_photo['Project']}. "
            if each_photo["Site"]:
                caption += f"Site: {each_photo['Site']}. "
            caption += each_photo["Description"]
            subprocess.run(
                [
                    "exiftool",
                    f"-artist={each_photo['Photographer']}",
                    f"-imagedescription={caption}",
                    "--usercomment",
                    Path(images_directory) / each_photo["Photo"],
                ]
            )
    main()


def view_csv_file() -> None:
    thefile = Path(images_directory) / "Photo Log.csv"
    try:
        os.startfile(thefile)
    except:
        os.system(f'open "{thefile}"')
    main()


def view_word_doc() -> None:
    thefile = Path(images_directory) / "Contact Sheet.docx"
    try:
        os.startfile(thefile)
    except:
        os.system(f'open "{thefile}"')
    main()


def main() -> None:
    global configs
    try:
        with open("configs.ini", "r") as f:
            pass
    except:
        with open("configs.ini", "w") as f:
            f.write("[DEFAULTS]\n")
            f.write("exiftool = /usr/local/bin/exiftool\n")
            f.write("# papersize options are 'a4' and 'letter'\n")
            f.write("papersize = a4\n")
            f.write("photographer = \n")
            f.write("project = \n")
            f.write("site = \n\n")
            f.write("[FACING]\n")
            f.write("# precision options are:\n")
            f.write("#  'coarse' (N, NE, E, SE, S, SW, W, NW)\n")
            f.write("#  'fine' (N, NNE, NE, ENE, E, and so-on)\n")
            f.write("#  'precise' (the actual bearing, in degrees)\n")
            f.write("precision = coarse\n")
    configs.read("configs.ini")
    exiftool = Path(configs["DEFAULTS"]["exiftool"])
    if not exiftool.is_file():
        exit(
            "EXITING:\nexiftool not found at the location indicated in the configs.ini file. Install it (https://exiftool.org) or update its path to use this script."
        )
    action = ""
    while action not in valid_actions:
        action = _display_menu()
    if action == "1":
        load_photos()
    elif action == "2":
        create_csv()
    elif action == "3":
        view_csv_file()
    elif action == "4":
        annotate_photos()
    elif action == "5":
        create_word_doc()
    elif action == "6":
        view_word_doc()
    elif action == "7":
        update_originals()
    elif action == "Q":
        exit()


if __name__ == "__main__":
    main()
