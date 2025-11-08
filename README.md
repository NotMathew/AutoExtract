**📋 Description**

An intelligent archive extraction tool that automatically processes compressed files across multiple formats with advanced password handling and file management capabilities.

**🎯 Key Features**

- **Multi-Format Support:** Handles ZIP, RAR, 7Z, TAR, GZ, BZ2, XZ, and compound archives
- **Multiple Scan Modes:** Choose between current directory only or recursive scanning through all subdirectories
- Smart Password Management: Three password policies:
  - Ask for password for each encrypted archive
  - Use same password for all archives
  - Skip all password-protected archives
- **Dual Extraction Engine:**
  - Primary: 7-Zip command line (most reliable)
  - Fallback: Patool library (broad format support)
- **Organized Extraction:** Creates dedicated folders for each archive, prevents overwrites
- **Post-Extraction Options:** Copy all extracted files to a single directory or selective copying
- **Cross-Platform:** Fully compatible with Windows and Linux systems
- **Detailed Reporting:** Comprehensive summary of extraction results and file statistics

**📁 Supported Platforms**

- Windows
- Linux
- macOS

**📦 Dependencies**

    ```
    patool
    rarfile
    ```

## Installation on Windows 10/11

Go to [Releases](https://github.com/NotMathew/AutoExtract.git) and download the lastest version.

## Installation on Linux

```
git clone https://github.com/NotMathew/AutoExtract.git
cd AutoExtract
sudo python -m pip install -r requirements.txt
python AutoExtract.py
```
