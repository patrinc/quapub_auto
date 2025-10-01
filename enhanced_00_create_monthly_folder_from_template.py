import shutil
import os
import sys
from pathlib import Path

# Add utils to path
utils_path = Path(__file__).parent / 'utils'
if str(utils_path) not in sys.path:
    sys.path.insert(0, str(utils_path))

# Import enhanced utilities
from utils import (
    setup_logger, ProcessingMetrics, ensure_directory_exists,
    BackupManager, get_config, validate_prerequisites
)

# Import enhanced config
try:
    import enhanced_config as config
    print("✅ Using enhanced configuration")
except ImportError:
    # Fallback to original values
    config = None
    print("⚠️  Using basic configuration")

class MonthlyFolderCreator:
    """Enhanced monthly folder creator with validation and logging"""
    
    def __init__(self):
        self.env_config = get_config()
        self.logger = setup_logger('folder_creator', self.env_config.log_level)
        self.metrics = ProcessingMetrics('folder_creator')
        
        # Configuration
        if config:
            self.base_path = config.BASE_PATHS['onedrive_base']
            self.month_folder_new = config.get_monthly_folder()
            self.template_folder = self.base_path.parent / "template_structure_with_scripts"
        else:
            # Fallback to original hardcoded values
            self.base_path = Path(r"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning")
            self.month_folder_new = "25_10 Monthly Scans"
            self.template_folder = self.base_path.parent / "template_structure_with_scripts"
        
        self.new_folder_path = self.base_path / self.month_folder_new
        
        self.logger.info("Monthly folder creator initialized")
        self.logger.info(f"Base path: {self.base_path}")
        self.logger.info(f"New folder: {self.month_folder_new}")
        self.logger.info(f"Template folder: {self.template_folder}")
    
    def validate_prerequisites(self) -> bool:
        """Validate that prerequisites exist"""
        with self.metrics.track_operation("validate_prerequisites"):
            # Check if template folder exists
            validation_result = validate_prerequisites(
                "folder_creator",
                required_paths=[self.template_folder],
                logger=self.logger
            )
            
            if not validation_result.valid:
                print(f"❌ Validation failed: {validation_result.get_summary()}")
                return False
            
            # Check if base path is writable
            if not os.access(self.base_path.parent, os.W_OK):
                self.logger.error(f"No write permission for base path: {self.base_path.parent}")
                print(f"❌ No write permission for: {self.base_path.parent}")
                return False
            
            print("✅ Prerequisites validation passed")
            return True
    
    def check_existing_folder(self) -> bool:
        """Check if folder already exists and handle accordingly"""
        with self.metrics.track_operation("check_existing") as op:
            if self.new_folder_path.exists():
                self.logger.warning(f"Folder already exists: {self.new_folder_path}")
                
                # Ask user what to do
                print(f"⚠️  Folder '{self.month_folder_new}' already exists!")
                print("Options:")
                print("1. Skip creation (folder exists)")
                print("2. Backup existing and recreate")
                print("3. Merge with existing (copy missing files)")
                print("4. Cancel operation")
                
                try:
                    choice = input("Enter your choice (1-4): ").strip()
                    
                    if choice == '1':
                        print("✅ Skipping - folder already exists")
                        op.details['action'] = 'skipped'
                        return False
                    elif choice == '2':
                        return self.backup_and_recreate()
                    elif choice == '3':
                        op.details['action'] = 'merge'
                        return True  # Proceed with merge
                    else:
                        print("❌ Operation cancelled")
                        op.details['action'] = 'cancelled'
                        return False
                        
                except KeyboardInterrupt:
                    print("\n❌ Operation cancelled by user")
                    return False
            
            op.details['action'] = 'create_new'
            return True
    
    def backup_and_recreate(self) -> bool:
        """Backup existing folder and recreate"""
        with self.metrics.track_operation("backup_and_recreate") as op:
            try:
                if self.env_config.create_backups:
                    backup_manager = BackupManager(logger=self.logger)
                    
                    # Create backup of existing folder
                    backup_path = backup_manager.backup_files([self.new_folder_path], "folder_recreation")
                    print(f"💾 Created backup: {backup_path}")
                    self.logger.info(f"Backed up existing folder to: {backup_path}")
                
                # Remove existing folder
                shutil.rmtree(self.new_folder_path)
                self.logger.info(f"Removed existing folder: {self.new_folder_path}")
                
                op.details['backup_created'] = True
                return True
                
            except Exception as e:
                self.logger.log_error_with_context(e, "backing up existing folder")
                print(f"❌ Failed to backup existing folder: {e}")
                op.details['error'] = str(e)
                return False
    
    def copy_template_structure(self) -> bool:
        """Copy template structure to new folder"""
        operation_name = "copy_template"
        
        with self.metrics.track_operation(operation_name) as op:
            try:
                self.logger.log_operation_start(operation_name, f"Copying template to {self.new_folder_path}")
                
                # Create the new folder if it doesn't exist
                ensure_directory_exists(self.new_folder_path)
                
                # Count files in template for progress tracking
                template_files = list(self.template_folder.rglob('*'))
                total_files = len([f for f in template_files if f.is_file()])
                
                self.logger.info(f"Found {total_files} files in template")
                
                # Set up progress tracking if enabled
                if self.env_config.enable_progress:
                    from utils.progress import ProgressTracker
                    progress = ProgressTracker(total_files, "Copying Files")
                
                # Copy with progress tracking
                def copy_with_progress(src, dst):
                    """Custom copy function with progress tracking"""
                    shutil.copytree(src, dst, dirs_exist_ok=True)
                    
                    if self.env_config.enable_progress:
                        # Update progress for all files (approximation)
                        progress.update(total_files, "Copy complete")
                
                # Perform the copy
                copy_with_progress(self.template_folder, self.new_folder_path)
                
                if self.env_config.enable_progress:
                    progress.complete()
                
                # Count copied files for verification
                copied_files = list(self.new_folder_path.rglob('*'))
                copied_count = len([f for f in copied_files if f.is_file()])
                
                op.details.update({
                    'template_files': total_files,
                    'copied_files': copied_count,
                    'success': True
                })
                
                self.metrics.increment_counter('files_copied', copied_count)
                
                self.logger.info(f"Successfully copied {copied_count} files")
                print(f"✅ Copied {copied_count} files from template")
                
                return True
                
            except Exception as e:
                self.logger.log_error_with_context(e, "copying template structure")
                print(f"❌ Failed to copy template: {e}")
                op.details.update({
                    'error': str(e),
                    'success': False
                })
                return False
    
    def update_config_files(self) -> bool:
        """Update config files in the new folder with current month/year"""
        with self.metrics.track_operation("update_configs") as op:
            try:
                if not config:
                    self.logger.info("No enhanced config available, skipping config updates")
                    return True
                
                # Find config files in the new folder
                config_files = list(self.new_folder_path.rglob('config.py'))
                
                if not config_files:
                    self.logger.warning("No config.py files found to update")
                    return True
                
                self.logger.info(f"Found {len(config_files)} config files to update")
                
                for config_file in config_files:
                    # Read current config
                    with open(config_file, 'r') as f:
                        content = f.read()
                    
                    # Update with current values
                    updated_content = content.replace(
                        f'fixed_start_date = date(2025, 9, 29)',
                        f'fixed_start_date = date({config.fixed_start_date.year}, {config.fixed_start_date.month}, {config.fixed_start_date.day})'
                    )
                    updated_content = updated_content.replace(
                        f'assessment_date = "29-09-25"',
                        f'assessment_date = "{config.assessment_date}"'
                    )
                    
                    # Write updated config
                    with open(config_file, 'w') as f:
                        f.write(updated_content)
                    
                    self.logger.info(f"Updated config file: {config_file}")
                
                op.details['config_files_updated'] = len(config_files)
                print(f"✅ Updated {len(config_files)} config files")
                
                return True
                
            except Exception as e:
                self.logger.log_error_with_context(e, "updating config files")
                print(f"⚠️  Warning: Failed to update config files: {e}")
                op.details['error'] = str(e)
                return True  # Don't fail the whole operation for this
    
    def create_monthly_folder(self) -> bool:
        """Main method to create monthly folder"""
        self.logger.info("Starting monthly folder creation")
        print(f"🚀 Creating monthly folder: {self.month_folder_new}")
        
        try:
            # Validate prerequisites
            if not self.validate_prerequisites():
                return False
            
            # Check existing folder
            if not self.check_existing_folder():
                return True  # Not an error, just skipped or cancelled
            
            # Copy template structure
            if not self.copy_template_structure():
                return False
            
            # Update config files
            self.update_config_files()
            
            # Complete metrics
            self.metrics.complete_script()
            
            # Print summary
            self.print_summary()
            
            print(f"✅ New monthly folder '{self.month_folder_new}' created successfully!")
            self.logger.info(f"Monthly folder creation completed successfully: {self.new_folder_path}")
            
            return True
            
        except Exception as e:
            self.logger.log_error_with_context(e, "creating monthly folder")
            print(f"❌ Failed to create monthly folder: {e}")
            return False
    
    def print_summary(self):
        """Print creation summary"""
        summary = self.metrics.get_summary()
        
        print("\n" + "="*60)
        print("📁 FOLDER CREATION SUMMARY")
        print("="*60)
        print(f"📂 Folder: {self.month_folder_new}")
        print(f"📍 Location: {self.new_folder_path}")
        print(f"⏱️  Total Time: {summary['total_time']:.2f}s")
        print(f"🔄 Operations: {summary['total_operations']}")
        
        if summary['counters']:
            print("📊 Stats:")
            for name, count in summary['counters'].items():
                print(f"   {name}: {count:,}")
        
        print("="*60)

def main():
    """Main entry point"""
    creator = MonthlyFolderCreator()
    
    try:
        success = creator.create_monthly_folder()
        sys.exit(0 if success else 1)
        
    except KeyboardInterrupt:
        print("\n❌ Operation cancelled by user")
        creator.logger.warning("Operation cancelled by user")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        creator.logger.error(f"Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 