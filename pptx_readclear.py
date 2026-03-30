from myfunction import process_directory
import argparse

def main():
    parser = argparse.ArgumentParser(description='Decrypt read-only .pptx files and copy other files in a directory.')
    parser.add_argument('source_path', help='Source file or directory containing the .pptx files.')
    parser.add_argument('--target_path', '-t', help='Target file or directory where decrypted files will be saved.', default=None)
    parser.add_argument('--in-place', '-i', action='store_true', help='Modify files in place (overwrites original files).')
    
    args = parser.parse_args()

    if not args.in_place and not args.target_path:
        parser.error('Either --target_path/-t or --in-place/-i must be provided.')

    process_directory(args.source_path, target_path=args.target_path, in_place=args.in_place)

if __name__ == '__main__':
    main()
