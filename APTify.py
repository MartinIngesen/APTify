from docx import Document
import argparse

parser = argparse.ArgumentParser(
    formatter_class=argparse.RawDescriptionHelpFormatter,
    epilog='''
SUPPORTED APTs
==============
APTs:
\t"GUCCIFER2.0" - Adds the name 'Феликс Эдмундович' ONLY to 'Last Modified By'
Generic:
\t"RUSSIA", - Adds the generic name 'Дмитрий' to 'Author' and 'Last Modified By'
''')

parser.add_argument("path", help="The document you want to APTify!")
parser.add_argument("group", help="The APT group you want to look like.")
args = parser.parse_args()

print("    _    ____ _____ _  __       ")
print("   / \  |  _ \_   _(_)/ _|_   _ ")
print("  / _ \ | |_) || | | | |_| | | |")
print(" / ___ \|  __/ | | | |  _| |_| |")
print("/_/   \_\_|    |_| |_|_|  \__, |")
print("   APTify your documents! |___/ ")
print("Created by Martin 'Mrtn' Ingesen")
print("="*32)
print("APTifying document...")

try:
    document = args.path
    group = args.group.lower()
    changed = False

    doc = Document(document)
    if(group == "russia"):
        print("Adding 'Дмитрий' to 'Author' and 'Last Modified By'")
        doc.core_properties.author = "Дмитрий"
        doc.core_properties.last_modified_by = "Дмитрий"
        changed = True
    elif(group == "guccifer2.0"):
        print("Adding 'Феликс Эдмундович' to 'Last Modified By'")
        doc.core_properties.last_modified_by = "Феликс Эдмундович"
        changed = True

    if changed:
        doc.save(document)
        print("APTification is complete.")
    else:
        print("No edits were made.")
except:
    print("APTification failed!")
