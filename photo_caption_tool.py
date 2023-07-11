import configparser
import csv
import math
import os
import re
from pathlib import Path
from datetime import datetime

from exif import Image as exifImage
from PIL import Image as pilImage, ImageOps, ImageFont, ImageDraw


valid_actions = []
images_directory = ""
all_images_exif_data = {}
rotation = [1, 8, 3, 6]  # Rotation of images, as represented in EXIF


def load_images_directory(precision: str) -> None:
    global images_directory
    images_directory = (
        input("Images Folder\n(type full path or drag folder onto here): ")
        .replace("\\", "")
        .strip("'")
        .strip()
    )
    images = []
    items = Path(images_directory).glob("*.*")
    for each_item in items:
        if re.match(".*\.jpe?g", each_item.name, flags=re.IGNORECASE):
            images.append(each_item.name)
            images.sort()
    if len(images) == 0:
        images_directory = ""
        print("NOTICE: There were no valid JPEG images in the selected directory.")
    create_csv(precision)
    main()


def facing(azimuth, precision) -> str:
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
    if precision == "fine":
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
        if precision == "precise":
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
    print("--------------------------")
    print(" Choose an action:")
    print(" 1 - Create New Photo Log")
    valid_actions = ["1"]
    if images_directory:
        print(" 2 - Annotate Photos")
        valid_actions.append("2")
        print(" 3 - Create Word Document")
        valid_actions.append("3")
    print(" Q - Quit")
    valid_actions.append("Q")
    print("--------------------------")
    return input().upper()


def create_csv(precision) -> None:
    global all_images_exif_data
    all_images_exif_data = {}
    if not images_directory:
        main()
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
    csv_file = Path(images_directory) / "Photo Log.csv"
    if csv_file.is_file():
        if (
            input(f"“{csv_file}” already exists. Type “Y” to overwrite it. ").upper()
            != "Y"
        ):
            main()
    images = []
    items = Path(images_directory).glob("*.*")
    for each_item in items:
        if re.match(".*\.jpe?g", each_item.name, flags=re.IGNORECASE):
            images.append(each_item.name)
            images.sort()
    if len(images) == 0:
        print("NOTICE: There were no valid JPEG images in the selected directory.")
    else:
        data_for_csv = []
        for each_image in images:
            image_data = {"Photo": each_image}
            with open(Path(images_directory) / each_image, "rb") as f:
                the_image = exifImage(f)
                all_images_exif_data[each_image] = the_image
                image_data["Photographer"] = the_image.get("artist")
                try:
                    timestamp = str(
                        datetime.strptime(
                            the_image.get("datetime_original"), "%Y:%m:%d %H:%M:%S"
                        )
                    )
                except:
                    timestamp = ""
                image_data["Timestamp"] = timestamp
                image_data["GPS Coordinates"] = parse_gps_coords(
                    the_image.get("gps_latitude"),
                    the_image.get("gps_longitude"),
                    the_image.get("gps_latitude_ref"),
                    the_image.get("gps_longitude_ref"),
                )
                image_data["Description"] = ""
                # Photo taken with iOS Camera.app:
                if the_image.get("image_description"):
                    image_data["Description"] = the_image.get(
                        "image_description"
                    ).strip()
                # Photo taken with Theodolite.app:
                if the_image.get("user_comment"):
                    image_data["Description"] = the_image.get("user_comment").strip()
                image_data["Facing"] = facing(
                    the_image.get("gps_img_direction"), precision
                )
            data_for_csv.append(image_data)
        with open(csv_file, "w") as f:
            csv_out = csv.DictWriter(f, fieldnames=headers)
            csv_out.writeheader()
            csv_out.writerows(data_for_csv)
        print("“Photo Log.csv” created.")
        if input("Type “Y” to open this file now. ").upper() == "Y":
            try:
                os.startfile(csv_file)
            except:
                os.system(f'open "{csv_file}"')
        main()


def annotate_photos(upateoriginals) -> None:
    global all_images_exif_data
    if not images_directory:
        main()
    output_dir = Path(images_directory) / "Annotated Photos"
    if output_dir.is_dir():
        if (
            input(f"“{output_dir}” already exists. Type “Y” to overwrite it. ").upper()
            != "Y"
        ):
            main()
    output_dir.mkdir(exist_ok=True)
    csv_file = Path(images_directory) / "Photo Log.csv"
    with open(csv_file, "r") as f:
        reader = csv.DictReader(f)
        for each_photo in reader:
            label = []
            line = []
            if each_photo["Project"]:
                line.append(f"Project: {each_photo['Project']}")
            if each_photo["Site"]:
                line.append(f"Site: {each_photo['Site']}")
            if line:
                label.append(".  ".join(line))
            if each_photo["Description"]:
                label.append(each_photo["Description"])
            line = []
            if each_photo["Facing"]:
                line.append(f"Facing {each_photo['Facing']}")
            if each_photo["GPS Coordinates"]:
                line.append(each_photo["GPS Coordinates"])
            if each_photo["Timestamp"]:
                line.append(each_photo["Timestamp"])
            if line:
                label.append(".  ".join(line))
            if not label:
                label = ["(unlabeled)"]
            orientation = all_images_exif_data[each_photo["Photo"]].get("orientation")
            img = pilImage.open(Path(images_directory) / each_photo["Photo"])
            img = img.rotate(rotation.index(orientation) * 90, expand=True)
            img = ImageOps.pad(img, (img.width, img.height + 200), centering=(0, 0))
            img_annotion = ImageDraw.Draw(img, mode="RGB")
            thefont = ImageFont.truetype("Helvetica.ttc", 46)
            img_annotion.text(
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
            update_exif_data(
                images_directory,
                each_photo["Photo"],
                filename,
                each_photo["Photographer"],
                each_photo["Description"],
            )
            # TODO: output a count of the total number of images annotated


def update_exif_data(
    images_directory: Path,
    original: str,
    annotated: str,
    photographer: str,
    description: str,
) -> None:
    # First update all_images_exif_data with the values from the CSV
    all_images_exif_data[original].set("artist", photographer)
    if all_images_exif_data[original].get("user_comment"):
        all_images_exif_data[original].set("user_comment", description[:-1])
    else:
        all_images_exif_data[original].set("image_description", description)
    # Next update the original photo.
    with open(Path(images_directory) / original, "wb") as f:
        f.write(all_images_exif_data[original].get_file())
    # Finally update the annotated photo.
    with open(Path(images_directory) / "Annotated Photos" / annotated, "rb") as f:
        all_images_exif_data[original].set("get_file", f.read())
    with open(Path(images_directory) / "Annotated Photos" / annotated, "wb") as f:
        all_images_exif_data[original].delete("pixel_x_dimension")
        all_images_exif_data[original].delete("pixel_y_dimension")
        f.write(all_images_exif_data[original].get_file)


def main() -> None:
    configs = configparser.ConfigParser()
    try:
        with open("configs.ini", "r") as f:
            pass
    except:
        with open("configs.ini", "w") as f:
            f.write("[FACING]\n")
            f.write("# options are:\n")
            f.write("#  'coarse' (N, NE, E, SE, S, SW, W, NW)\n")
            f.write("#  'fine' (N, NNE, NE, ENE, E, and so-on)\n")
            f.write("#  'precise' (the actual bearing, in degrees)\n")
            f.write("precision = coarse\n\n")
            f.write("[CAPTIONS]\n")
            f.write("# write the edited captions back to the original images\n")
            f.write("updateoriginals = yes")
    configs.read("configs.ini")
    action = ""
    while action not in valid_actions:
        action = display_menu()
    if action == "1":
        load_images_directory(configs["FACING"]["precision"])
    elif action == "2":
        annotate_photos(configs["CAPTIONS"]["updateoriginals"])
    elif action == "Q":
        exit()


if __name__ == "__main__":
    main()

# TODO: clear extraneous input that causes the annotation to run twice if the user bungled the "overwrite" warning
# TODO: create Word doc
# TODO: create config file that controls: coarse/fine/precise facing, whether to write corrected captions back to originals, ??
