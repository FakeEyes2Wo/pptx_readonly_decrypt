from myfunction import process_directory
import argparse

def main():
    parser = argparse.ArgumentParser(description='Decrypt read-only .pptx files and copy other files in a directory.')
    parser.add_argument('--source_path', help='Source directory containing the .pptx files and other files.')
    parser.add_argument('--target_path', help='Target directory where decrypted .pptx files and other files will be saved.')
    
    args = parser.parse_args()

    process_directory(args.source_path, args.target_path)

if __name__ == '__main__':
    main()

##python pptx_readclear.py source_dir target_dir
