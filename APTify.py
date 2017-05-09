from docx import Document
import argparse

parser = argparse.ArgumentParser()

parser.add_argument("path", help="The document you want to APTify!")
args = parser.parse_args()

print("     _    ____ _____ _  __       ")
print("    / \  |  _ \_   _(_)/ _|_   _ ")
print("   / _ \ | |_) || | | | |_| | | |")
print("  / ___ \|  __/ | | | |  _| |_| |")
print(" /_/   \_\_|    |_| |_|_|  \__, |")
print("    APTify your documents! |___/ ")
print("Created by Martin 'Mrtn' Ingesen ")
print("="*33)
print("APTifying document...")
try:
    document = args.path
    doc = Document(document)
    doc.core_properties.author = "Дмитрий"
    doc.save(document)
    print("APTification is complete.")
except:
    print("APTification failed!")
