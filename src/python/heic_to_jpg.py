import os
import sys
from PIL import Image
from pillow_heif import register_heif_opener

register_heif_opener()


def convert_heic_to_jpg(input_path, output_path=None):
    if not output_path:
        base_name = os.path.splitext(input_path)[0]
        output_path = f"{base_name}.jpg"

    try:
        image = Image.open(input_path)
        image.save(output_path, "JPEG", quality=95, exif=image.getexif())
        print(f"Success: HEIC file converted to {output_path}")
        return 0

    except FileNotFoundError:
        print(f"Error: File not found ({input_path})")
        return 2

    except Exception as e:
        print(f"Error: Conversion failed for {input_path} - {e}")
        return 1


def main():
    if len(sys.argv) < 3:
        print("Usage: heic_to_jpg.exe <input_heic_path> <output_jpg_path>")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    result_code = convert_heic_to_jpg(input_path, output_path)

    if result_code != 0:
        sys.exit(result_code)


if __name__ == "__main__":
    main()