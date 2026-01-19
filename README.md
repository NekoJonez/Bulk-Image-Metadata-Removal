## Bulk Image Metadata Removal

This PowerShell script allows you to **remove metadata from images in bulk**, while leaving the actual image data untouched.

It is designed for creators, photographers, AI users, and privacy-conscious users who want a **simple, transparent way** to clean image metadata before sharing files.

---

## What this script does

* Removes common metadata such as:

  * EXIF data
  * camera and device information
  * software tags
  * GPS/location data (if present)
* Works on **multiple images at once**
* Processes files **locally** on your machine
* Does **not** upload images anywhere
* Does **not** modify pixel data or image quality

Only metadata is affected.

---

## What this script does *not* do

* It does **not** compress or resize images
* It does **not** alter colors or image content
* It does **not** require internet access (if ExifTool is installed)
* It does **not** rely on external services or APIs

This is a local, offline cleanup tool.

---

## Why this exists

Image metadata is often invisible, but it can contain more information than intended when sharing files publicly.

This tool exists to:

* give users control over their images
* reduce accidental data leakage
* provide a readable, auditable alternative to closed tools

You can inspect the script yourself to see exactly what it does.

---

## Usage

1. Download or clone the repository
2. Run the script in PowerShell
3. Select the folder containing your images
4. The script processes supported image files and removes metadata

*(Administrator rights are not required for normal use.)*

Or run the .exe file, which is basically the same but you don't have to execute Powershell. 
The manual is located in the repo as well.

---

## Safety & transparency

* The script is **non-destructive** to image content
* Original image structure is preserved
* All operations are performed locally
* Source code is fully visible and modifiable

If you want to test first, use it on a copy of your images.

---

## License

This project is released for public use.
See https://github.com/NekoJonez/Bulk-Image-Metadata-Removal/tree/main?tab=BSD-3-Clause-1-ov-file#

— Made with care (and a little floof) by NekoJonez 🐾
