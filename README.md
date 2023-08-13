# Automation Script

This script automates the organization of files and updating of applications. It monitors the Downloads folder for file events and sorts files into appropriate folders based on their types. Additionally, it performs app updates using the `winget` command once a week.

### This is what it does so far:
- Automatically organizes files into designated directories based on their types (images, documents, installations).
- Identifies .part files and employs a brief delay of 10 seconds to ensure that downloads are complete before proceeding with moving .exe or.msi files
- Extracts contents from ZIP files, places them in an “Extracted Zip Files” folder, and deletes the empty ZIP files.
- Updates installed applications using the `winget` command once a week.

I'm hoping to add more functionality in the future.
