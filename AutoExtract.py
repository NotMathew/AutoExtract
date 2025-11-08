import os
import zipfile
import shutil
from pathlib import Path
import sys
from collections import defaultdict
import subprocess
import platform

class ArchiveExtractor:
    def __init__(self):
        self.supported_formats = {'.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.xz', '.tar.gz', '.tar.bz2', '.tar.xz'}
        self.extraction_results = defaultdict(list)
        self.total_archives = 0
        self.successful_extractions = 0
        self.failed_extractions = 0
        self.password_protected = 0
        self.global_password = None
        self.all_extracted_files = []
        
    def format_size(self, size_bytes):
        """Convert bytes to human readable format"""
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} PB"
    
    def get_archive_size(self, filepath):
        """Get the size of the archive file"""
        try:
            return os.path.getsize(filepath)
        except OSError:
            return 0
    
    def find_archives(self, directory, recursive=True):
        """Find all supported archive files in directory"""
        archives = []
        print(f"🔍 Searching for archive files in: {directory}")
        
        if recursive:
            for root, dirs, files in os.walk(directory):
                for file in files:
                    filepath = os.path.join(root, file)
                    if self.is_supported_archive(filepath):
                        archives.append(filepath)
        else:
            for item in os.listdir(directory):
                filepath = os.path.join(directory, item)
                if os.path.isfile(filepath) and self.is_supported_archive(filepath):
                    archives.append(filepath)
        
        return archives
    
    def is_supported_archive(self, filepath):
        """Check if file is a supported archive format"""
        extension = Path(filepath).suffix.lower()
        return extension in self.supported_formats
    
    def ask_scan_mode(self):
        """Ask user whether to scan current directory only or include subdirectories"""
        print("\n📁 Scan Mode Selection")
        print("1. Current directory only (no subdirectories)")
        print("2. Current directory and all subdirectories (recursive)")
        
        while True:
            choice = input("\nEnter your choice (1-2): ").strip()
            if choice == '1':
                return False
            elif choice == '2':
                return True
            else:
                print("Invalid choice! Please enter 1 or 2.")
    
    def ask_password_policy(self):
        """Ask user about password handling policy"""
        print("\n🔐 Password Handling")
        print("If encrypted archives are found, how should passwords be handled?")
        print("1. Ask for password for each encrypted archive")
        print("2. Use same password for all encrypted archives")
        print("3. Skip all encrypted archives")
        
        while True:
            choice = input("\nEnter your choice (1-3): ").strip()
            if choice == '1':
                return 'ask_each'
            elif choice == '2':
                self.global_password = input("Enter the password to use for all archives: ").strip()
                return 'use_global'
            elif choice == '3':
                return 'skip_all'
            else:
                print("Invalid choice! Please enter 1, 2, or 3.")
    
    def get_extraction_path(self, archive_path):
        """Generate extraction path based on archive name"""
        archive_dir = os.path.dirname(archive_path)
        archive_name = Path(archive_path).stem
        extraction_dir = os.path.join(archive_dir, archive_name)
        
        counter = 1
        original_dir = extraction_dir
        while os.path.exists(extraction_dir):
            extraction_dir = f"{original_dir}_{counter}"
            counter += 1
        
        return extraction_dir
    
    def collect_extracted_files(self, extraction_path):
        """Collect all files from extraction directory"""
        extracted_files = []
        total_size = 0
        
        try:
            for root, dirs, files in os.walk(extraction_path):
                for file in files:
                    filepath = os.path.join(root, file)
                    try:
                        file_size = os.path.getsize(filepath)
                        extracted_files.append({
                            'path': filepath,
                            'size': file_size,
                            'relative_path': os.path.relpath(filepath, extraction_path)
                        })
                        total_size += file_size
                    except OSError:
                        continue
        except OSError:
            pass
        
        return extracted_files, total_size
    
    def extract_with_patool(self, archive_path, extraction_path, password=None):
        """Extract archive using patool (handles most formats)"""
        try:
            import patoolib
            # Create extraction directory
            os.makedirs(extraction_path, exist_ok=True)
            
            if password:
                # Try with password first
                try:
                    patoolib.extract_archive(archive_path, outdir=extraction_path, password=password)
                    return True, "Success"
                except patoolib.util.PatoolError:
                    # If password fails, try without (some archives might not need it)
                    try:
                        patoolib.extract_archive(archive_path, outdir=extraction_path)
                        return True, "Success"
                    except patoolib.util.PatoolError as e:
                        error_msg = str(e)
                        if 'password' in error_msg.lower() or 'encrypted' in error_msg.lower():
                            self.password_protected += 1
                            return False, "Password required or incorrect password"
                        else:
                            return False, error_msg
            else:
                patoolib.extract_archive(archive_path, outdir=extraction_path)
                return True, "Success"
                
        except patoolib.util.PatoolError as e:
            error_msg = str(e)
            if 'password' in error_msg.lower() or 'encrypted' in error_msg.lower():
                self.password_protected += 1
                return False, "Password required or incorrect password"
            else:
                return False, error_msg
        except Exception as e:
            return False, str(e)
    
    def get_7zip_paths(self):
        """Get 7-Zip executable paths for current platform"""
        if platform.system() == "Windows":
            return [
                "C:\\Program Files\\7-Zip\\7z.exe",
                "C:\\Program Files (x86)\\7-Zip\\7z.exe",
                "7z.exe",
                "7z"
            ]
        else:  # Linux, macOS, etc.
            return [
                "/usr/bin/7z",
                "/usr/bin/7za",
                "/usr/local/bin/7z",
                "/usr/local/bin/7za",
                "7z",
                "7za"
            ]
    
    def find_7zip_executable(self):
        """Find 7-Zip executable in system"""
        seven_zip_paths = self.get_7zip_paths()
        
        for path in seven_zip_paths:
            # Check if path exists
            if os.path.exists(path):
                return path
            
            # Check if command is in PATH
            try:
                if platform.system() == "Windows":
                    result = subprocess.run([path, "--help"], capture_output=True, timeout=5, check=False)
                else:
                    result = subprocess.run([path, "--help"], capture_output=True, timeout=5, check=False)
                
                if result.returncode in [0, 1]:  # 7z returns 1 for --help on some systems
                    return path
            except (subprocess.TimeoutExpired, FileNotFoundError, PermissionError):
                continue
        
        return None
    
    def extract_with_7zip(self, archive_path, extraction_path, password=None):
        """Extract using 7-Zip command line (most reliable)"""
        try:
            seven_zip_exe = self.find_7zip_executable()
            
            if not seven_zip_exe:
                return False, "7-Zip not found. Please install 7-Zip/p7zip"
            
            # Build 7-Zip command
            if platform.system() == "Windows":
                cmd = [seven_zip_exe, "x", archive_path, f"-o{extraction_path}", "-y"]
            else:
                cmd = [seven_zip_exe, "x", archive_path, f"-o{extraction_path}", "-y"]
            
            if password:
                cmd.extend([f"-p{password}"])
            
            # Run 7-Zip
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            
            if result.returncode == 0:
                return True, "Success"
            else:
                error_output = result.stderr.lower() + result.stdout.lower()
                if "wrong password" in error_output or "encrypted" in error_output:
                    self.password_protected += 1
                    return False, "Wrong password or encrypted archive"
                elif "not supported" in error_output:
                    return False, "Compression method not supported"
                else:
                    return False, f"7-Zip error (code {result.returncode}): {result.stderr}"
                    
        except subprocess.TimeoutExpired:
            return False, "Extraction timed out"
        except Exception as e:
            return False, str(e)
    
    def extract_archive(self, archive_path, password_policy, current_password=None):
        """Extract a single archive file"""
        archive_name = os.path.basename(archive_path)
        extraction_path = self.get_extraction_path(archive_path)
        archive_size = self.get_archive_size(archive_path)
        
        print(f"\n📦 Extracting: {archive_name} ({self.format_size(archive_size)})")
        print(f"   Destination: {extraction_path}")
        
        # Handle password
        password = current_password
        max_attempts = 3
        attempts = 0
        
        while attempts < max_attempts:
            # If we don't have a password and policy is ask_each, ask for it
            if password_policy == 'ask_each' and not password:
                use_password = input(f"Does '{archive_name}' require a password? (y/N): ").strip().lower()
                if use_password == 'y':
                    password = input(f"Enter password for '{archive_name}': ").strip()
                    self.password_protected += 1
            
            # Try extraction methods in order of reliability
            success = False
            message = "Extraction failed"
            
            # Method 1: Try 7-Zip first (most reliable)
            success, message = self.extract_with_7zip(archive_path, extraction_path, password)
            
            if not success and "7-Zip not found" in message:
                # Method 2: Fall back to patool if 7-Zip not available
                success, message = self.extract_with_patool(archive_path, extraction_path, password)
            
            if success:
                break
            elif "password" in message.lower() and password_policy == 'ask_each' and attempts < max_attempts - 1:
                print(f"   ❌ Failed: {message}")
                password = input(f"Enter password for '{archive_name}' (attempt {attempts + 2}/{max_attempts}): ").strip()
                attempts += 1
                continue
            else:
                break
        
        # Clean up empty extraction directory if extraction failed
        if not success:
            try:
                if os.path.exists(extraction_path) and not os.listdir(extraction_path):
                    os.rmdir(extraction_path)
            except OSError:
                pass
        
        # Record result
        result = {
            'path': archive_path,
            'success': success,
            'message': message,
            'size': archive_size,
            'extraction_path': extraction_path if success else None
        }
        
        if success:
            extracted_files, total_size = self.collect_extracted_files(extraction_path)
            result['extracted_files'] = extracted_files
            result['extracted_size'] = total_size
            result['file_count'] = len(extracted_files)
            self.all_extracted_files.extend(extracted_files)
            self.successful_extractions += 1
            print(f"   ✅ Success: {message}")
            print(f"   📁 Extracted {len(extracted_files)} files ({self.format_size(total_size)})")
            self.extraction_results['success'].append(result)
        else:
            self.failed_extractions += 1
            print(f"   ❌ Failed: {message}")
            self.extraction_results['failed'].append(result)
        
        self.extraction_results['all'].append(result)
        return success
    
    def extract_all_archives(self, directory, recursive=True, password_policy='ask_each'):
        """Extract all archives in the directory"""
        archives = self.find_archives(directory, recursive)
        self.total_archives = len(archives)
        
        if not archives:
            print("❌ No supported archive files found!")
            return
        
        print(f"\n🎯 Found {self.total_archives} archive files")
        print("Starting extraction...")
        print("=" * 60)
        
        current_password = self.global_password if password_policy == 'use_global' else None
        
        for i, archive_path in enumerate(archives, 1):
            print(f"\n[{i}/{self.total_archives}] ", end="")
            self.extract_archive(archive_path, password_policy, current_password)
    
    def ask_copy_files(self):
        """Ask user if they want to copy extracted files to a specific path"""
        if not self.all_extracted_files:
            print("\nNo extracted files to copy.")
            return False
        
        total_files = len(self.all_extracted_files)
        total_size = sum(file['size'] for file in self.all_extracted_files)
        
        print(f"\n📋 Extraction completed! Found {total_files} files ({self.format_size(total_size)})")
        print("\nDo you want to copy all extracted files to a specific directory?")
        print("1. Yes, copy all files to a single directory")
        print("2. No, keep files in their original extraction folders")
        print("3. Yes, but let me choose which files to copy")
        
        while True:
            choice = input("\nEnter your choice (1-3): ").strip()
            
            if choice == '1':
                return self.copy_all_files()
            elif choice == '2':
                print("Files will remain in their extraction folders.")
                return False
            elif choice == '3':
                return self.selective_copy()
            else:
                print("Invalid choice! Please enter 1, 2, or 3.")
    
    def copy_all_files(self):
        """Copy all extracted files to a user-specified directory"""
        target_dir = input("\nEnter the target directory path: ").strip()
        
        # Validate target directory
        if not os.path.exists(target_dir):
            create_dir = input(f"Directory '{target_dir}' doesn't exist. Create it? (y/N): ").strip().lower()
            if create_dir == 'y':
                try:
                    os.makedirs(target_dir, exist_ok=True)
                    print(f"✅ Created directory: {target_dir}")
                except OSError as e:
                    print(f"❌ Failed to create directory: {e}")
                    return False
            else:
                print("Copy operation cancelled.")
                return False
        
        if not os.path.isdir(target_dir):
            print(f"❌ '{target_dir}' is not a directory!")
            return False
        
        print(f"\n📤 Copying {len(self.all_extracted_files)} files to: {target_dir}")
        print("=" * 60)
        
        copied_files = 0
        copied_size = 0
        skipped_files = 0
        
        for i, file_info in enumerate(self.all_extracted_files, 1):
            source_path = file_info['path']
            filename = os.path.basename(source_path)
            target_path = os.path.join(target_dir, filename)
            
            # Handle duplicate filenames
            counter = 1
            original_target = target_path
            while os.path.exists(target_path):
                name, ext = os.path.splitext(filename)
                target_path = os.path.join(target_dir, f"{name}_{counter}{ext}")
                counter += 1
            
            try:
                shutil.copy2(source_path, target_path)
                copied_files += 1
                copied_size += file_info['size']
                print(f"[{i}/{len(self.all_extracted_files)}] ✅ Copied: {filename}")
            except Exception as e:
                skipped_files += 1
                print(f"[{i}/{len(self.all_extracted_files)}] ❌ Failed: {filename} - {e}")
        
        # Show copy summary
        print(f"\n{'='*60}")
        print("COPY SUMMARY")
        print(f"{'='*60}")
        print(f"Total files attempted:  {len(self.all_extracted_files)}")
        print(f"Successfully copied:    {copied_files}")
        print(f"Failed/Skipped:         {skipped_files}")
        print(f"Total size copied:      {self.format_size(copied_size)}")
        print(f"Target directory:       {target_dir}")
        print(f"{'='*60}")
        
        return True
    
    def selective_copy(self):
        """Let user select which files to copy"""
        print(f"\n📋 Select files to copy ({len(self.all_extracted_files)} files found)")
        print("Enter file numbers separated by commas (e.g., 1,3,5) or 'all' for all files")
        
        # Display files with numbers
        for i, file_info in enumerate(self.all_extracted_files, 1):
            filename = os.path.basename(file_info['path'])
            print(f"  {i}. {filename} ({self.format_size(file_info['size'])})")
        
        while True:
            selection = input("\nEnter your selection: ").strip().lower()
            
            if selection == 'all':
                selected_files = self.all_extracted_files
                break
            else:
                try:
                    indices = [int(idx.strip()) - 1 for idx in selection.split(',')]
                    selected_files = [self.all_extracted_files[i] for i in indices if 0 <= i < len(self.all_extracted_files)]
                    if selected_files:
                        break
                    else:
                        print("No valid files selected. Please try again.")
                except ValueError:
                    print("Invalid input. Please enter numbers separated by commas or 'all'.")
        
        if not selected_files:
            print("No files selected for copying.")
            return False
        
        target_dir = input("\nEnter the target directory path: ").strip()
        
        # Validate target directory
        if not os.path.exists(target_dir):
            create_dir = input(f"Directory '{target_dir}' doesn't exist. Create it? (y/N): ").strip().lower()
            if create_dir == 'y':
                try:
                    os.makedirs(target_dir, exist_ok=True)
                    print(f"✅ Created directory: {target_dir}")
                except OSError as e:
                    print(f"❌ Failed to create directory: {e}")
                    return False
            else:
                print("Copy operation cancelled.")
                return False
        
        print(f"\n📤 Copying {len(selected_files)} files to: {target_dir}")
        print("=" * 60)
        
        copied_files = 0
        copied_size = 0
        
        for i, file_info in enumerate(selected_files, 1):
            source_path = file_info['path']
            filename = os.path.basename(source_path)
            target_path = os.path.join(target_dir, filename)
            
            # Handle duplicate filenames
            counter = 1
            original_target = target_path
            while os.path.exists(target_path):
                name, ext = os.path.splitext(filename)
                target_path = os.path.join(target_dir, f"{name}_{counter}{ext}")
                counter += 1
            
            try:
                shutil.copy2(source_path, target_path)
                copied_files += 1
                copied_size += file_info['size']
                print(f"[{i}/{len(selected_files)}] ✅ Copied: {filename}")
            except Exception as e:
                print(f"[{i}/{len(selected_files)}] ❌ Failed: {filename} - {e}")
        
        print(f"\n✅ Successfully copied {copied_files}/{len(selected_files)} files")
        print(f"📦 Total size: {self.format_size(copied_size)}")
        
        return True
    
    def show_summary(self):
        """Display comprehensive extraction summary"""
        print(f"\n{'='*80}")
        print("EXTRACTION SUMMARY")
        print(f"{'='*80}")
        
        print(f"Total archives found:     {self.total_archives}")
        print(f"Successfully extracted:   {self.successful_extractions}")
        print(f"Failed extractions:       {self.failed_extractions}")
        print(f"Password protected:       {self.password_protected}")
        
        if self.successful_extractions > 0:
            total_archive_size = sum(item['size'] for item in self.extraction_results['success'])
            total_extracted_size = sum(item['extracted_size'] for item in self.extraction_results['success'])
            total_files = sum(item['file_count'] for item in self.extraction_results['success'])
            
            print(f"Total archives size:      {self.format_size(total_archive_size)}")
            print(f"Total extracted size:     {self.format_size(total_extracted_size)}")
            print(f"Total files extracted:    {total_files}")
        
        success_rate = (self.successful_extractions / self.total_archives * 100) if self.total_archives > 0 else 0
        print(f"Success rate:             {success_rate:.1f}%")
        
        # Show failed extractions if any
        if self.extraction_results['failed']:
            print(f"\n❌ Failed extractions ({self.failed_extractions}):")
            for failed in self.extraction_results['failed']:
                print(f"   - {os.path.basename(failed['path'])}: {failed['message']}")
        
        # Show successful extractions if any
        if self.extraction_results['success']:
            print(f"\n✅ Successful extractions ({self.successful_extractions}):")
            for success in self.extraction_results['success']:
                file_count = success.get('file_count', 0)
                print(f"   - {os.path.basename(success['path'])} → {file_count} files")
        
        print(f"{'='*80}")
        
        if self.successful_extractions > 0:
            print("🎉 Extraction completed successfully!")
        else:
            print("ℹ️  No files were extracted.")

def main():
    print("📦 Archive File Extractor")
    print("=" * 50)
    print("Supported formats: ZIP, RAR, 7Z, TAR, GZ, BZ2, XZ")
    print(f"Platform: {platform.system()}")
    print("Using: 7-Zip (recommended) + patool (fallback)")
    print("=" * 50)
    
    # Get current directory
    current_dir = os.getcwd()
    print(f"Current directory: {current_dir}")
    
    extractor = ArchiveExtractor()
    
    try:
        # Ask for scan mode
        recursive = extractor.ask_scan_mode()
        scan_mode = "current directory and all subdirectories" if recursive else "current directory only"
        print(f"\nSelected scan mode: {scan_mode}")
        
        # Ask for password policy
        password_policy = extractor.ask_password_policy()
        
        # Extract archives
        extractor.extract_all_archives(current_dir, recursive, password_policy)
        
        # Ask about copying files
        extractor.ask_copy_files()
        
        # Show summary
        extractor.show_summary()
        
    except KeyboardInterrupt:
        print("\n\nProcess interrupted by user.")
        extractor.show_summary()
    except Exception as e:
        print(f"\n❌ Error: {e}")
        sys.exit(1)

# Check if required libraries are installed
def check_dependencies():
    missing = []
    
    try:
        import patoolib
    except ImportError:
        missing.append("patool")
    
    if missing:
        print("❌ Missing required libraries. Please install them using:")
        print("pip install patool")
        return False
    
    return True

if __name__ == "__main__":
    if check_dependencies():
        main()