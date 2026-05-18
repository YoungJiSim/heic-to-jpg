import os
import sys
import glob

from heic_to_jpg import convert_heic_to_jpg

def batch_conversion(input_dir, output_dir=None):
    if not input_dir or input_dir.strip() == "":
        input_dir = "."

    print(f"Batch conversion target folder: {os.path.abspath(input_dir)}")

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        print(f"Converted files destination: {os.path.abspath(output_dir)}")

    heic_files = glob.glob(os.path.join(input_dir, "*.heic")) + glob.glob(os.path.join(input_dir, "*.HEIC"))

    if not heic_files:
        print(f"Notice: No HEIC files found")
        sys.exit(2)

    print(f"{len(heic_files)} HEIC files found")
    print("Starting conversion...")

    count = 0
    for file_path in heic_files:

        if output_dir:
            file_name_with_extension = os.path.basename(file_path)
            file_name = os.path.splitext(file_name_with_extension)[0]
            target_output_path = os.path.join(output_dir, f"{file_name}.jpg")
        else:
            target_output_path = None

        result_code = convert_heic_to_jpg(file_path, target_output_path)

        if result_code == 0:
            count += 1

    print(f"{count} / {len(heic_files)} files converted to JPG")


def main():
    input_dir = sys.argv[1] if len(sys.argv) > 1 else "."
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    batch_conversion(input_dir, output_dir)


if __name__ == "__main__":
    main()