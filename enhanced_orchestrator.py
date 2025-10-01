import os
import re
import subprocess
import sys
from pathlib import Path
from colorama import init, Fore
from typing import List, Tuple, Optional, Dict, Any
import time

# Add utils to path if not already there
utils_path = Path(__file__).parent / 'utils'
if str(utils_path) not in sys.path:
    sys.path.insert(0, str(utils_path))

# Import enhanced utilities
from utils import (
    setup_logger, ProcessingMetrics, ValidationResult, 
    validate_prerequisites, BackupManager, get_config,
    print_environment_info, should_create_backups,
    should_run_validation, should_show_progress, is_dry_run
)
import custom_parsers

# Import enhanced config
try:
    import enhanced_config as config
    print("✅ Using enhanced configuration")
except ImportError:
    import config
    print("⚠️  Using basic configuration - consider upgrading")

init(autoreset=True)

class EnhancedOrchestrator:
    """Enhanced script orchestrator with validation, backups, and metrics"""
    
    def __init__(self):
        self.env_config = get_config()
        self.logger = setup_logger('orchestrator', self.env_config.log_level)
        self.backup_manager = BackupManager(
            max_backups=self.env_config.max_backups,
            logger=self.logger
        ) if should_create_backups() else None
        self.metrics = ProcessingMetrics('orchestrator')
        
        # Script parser mapping
        self.script_parsers = {
            "00_raw_report_network_segregation.py": custom_parsers.parse_network_segregation,
            "00_raw_report_server_segregation.py": custom_parsers.parse_server_segregation,
            "01_get_all_reports_from_sharepoint_latest.py": custom_parsers.parse_sharepoint_report,
            "02_delete_overview.py": custom_parsers.parse_delete_overview,
            "03_status_color.py": custom_parsers.parse_status_color,
            "04_exception_list.py": custom_parsers.parse_exception_list,
            "05_change_name_reports_to_report.py": custom_parsers.parse_change_name_reports,
            "06_process_and_update_reports.py": custom_parsers.parse_process_and_update_reports,
            "07_exception_removal.py": custom_parsers.parse_exception_removal,
            "08_copy_to_without_overview_folder.py": custom_parsers.parse_copy_without_overview,
            "09_insert_metadata.py": custom_parsers.parse_insert_metadata,
            "10_create_individual_overviews_with_date.py": custom_parsers.parse_create_individual_overviews,
            "11_generate_summary.py": custom_parsers.parse_generate_summary,
            "12_rename.py": custom_parsers.parse_rename,
            "13_delete_status.py": getattr(custom_parsers, "parse_delete_status", self._fallback_parser),
            "14_copy_to_sla_status.py": custom_parsers.parse_copy_to_sla_status,
            "15_overdue_calc.py": custom_parsers.parse_overdue_calc,
            "16_evidence_folder.py": custom_parsers.parse_evidence_folder,
            "17_old_new_status.py": custom_parsers.parse_old_new_status,
            "18_upload_final_reports_with_status_internal_sharepoint.py": custom_parsers.parse_upload_final_reports_sharepoint,
            "20_summary_ondemand.py": custom_parsers.parse_summary_ondemand,
            "21_export_charts.py": custom_parsers.parse_export_charts,
            "22_draft_internal_email.py": custom_parsers.parse_draft_internal_email,
            "23_reply_vapt_emails.py": custom_parsers.parse_reply_vapt_emails,
        }
    
    def _fallback_parser(self, stdout: str, stderr: str):
        """Fallback parser for scripts without custom parsers"""
        print(Fore.CYAN + "\n=== Script Output ===")
        if stdout:
            print(Fore.RESET + stdout)
        if stderr:
            print(Fore.RED + "Errors:\n" + stderr)
    
    def get_latest_monthly_folder(self, parent_dir: str = ".") -> Optional[str]:
        """Find the latest folder matching pattern: 'YY_MM Monthly Scans'"""
        with self.metrics.track_operation("find_latest_folder"):
            pattern = re.compile(r"(\d{2})_(\d{2}) Monthly Scans")
            folders = [f for f in os.listdir(parent_dir) if os.path.isdir(os.path.join(parent_dir, f))]
            matching_folders = []
            
            for folder in folders:
                match = pattern.match(folder)
                if match:
                    year, month = int(match.group(1)), int(match.group(2))
                    matching_folders.append((year, month, folder))
            
            if not matching_folders:
                self.logger.error("No matching monthly scan folders found")
                print(Fore.RED + "No matching monthly scan folders found.")
                return None
            
            # Sort by year and month descending to get latest
            matching_folders.sort(key=lambda x: (x[0], x[1]), reverse=True)
            latest_folder = matching_folders[0][2]
            folder_path = os.path.join(parent_dir, latest_folder)
            
            self.logger.info(f"Found latest monthly folder: {latest_folder}")
            return folder_path
    
    def extract_number(self, filename: str) -> int:
        """Extract leading number from filename for sorting"""
        match = re.match(r"(\d+)", filename)
        return int(match.group(1)) if match else float('inf')
    
    def gather_scripts(self, folder: str) -> List[Tuple[str, str]]:
        """Recursively gather all Python scripts in folder"""
        with self.metrics.track_operation("gather_scripts") as op:
            scripts = []
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.endswith(".py"):
                        full_path = os.path.join(root, file)
                        rel_path = os.path.relpath(full_path, folder)
                        scripts.append((rel_path, full_path))
            
            op.details['scripts_found'] = len(scripts)
            self.logger.info(f"Found {len(scripts)} Python scripts")
            return scripts
    
    def validate_script_prerequisites(self, script_name: str, script_dir: str) -> ValidationResult:
        """Validate prerequisites for a specific script"""
        if not should_run_validation():
            return ValidationResult(True, [], [], [], [])
        
        # Get required paths for this script
        if hasattr(config, 'get_required_paths_for_script'):
            required_paths = config.get_required_paths_for_script(script_name)
        else:
            # Fallback to basic validation
            required_paths = [script_dir]
        
        return validate_prerequisites(script_name, required_paths, logger=self.logger)
    
    def create_script_backup(self, script_path: str) -> Optional[str]:
        """Create backup before running script if needed"""
        if not should_create_backups() or not self.backup_manager:
            return None
        
        try:
            # Look for files this script might modify
            script_dir = Path(script_path).parent
            potential_files = list(script_dir.glob("*.xlsx")) + list(script_dir.glob("*.csv"))
            
            if potential_files:
                backup_paths = self.backup_manager.backup_files(
                    potential_files[:5],  # Limit to 5 files to avoid too many backups
                    category="pre_script"
                )
                return f"Created {len([b for b in backup_paths if b])} backups"
            
        except Exception as e:
            self.logger.error(f"Failed to create backups: {e}")
        
        return None
    
    def run_script(self, rel_path: str, full_path: str) -> bool:
        """Run a single script with enhanced features"""
        script_name = os.path.basename(full_path)
        script_dir = os.path.dirname(full_path)
        
        with self.metrics.track_operation(f"run_{script_name}") as op:
            self.logger.log_operation_start(f"Running script", script_name)
            
            # Validation
            if should_run_validation():
                validation_result = self.validate_script_prerequisites(script_name, script_dir)
                if validation_result.has_errors():
                    print(Fore.RED + f"❌ Validation failed for {script_name}")
                    print(Fore.RED + validation_result.get_summary())
                    self.logger.error(f"Validation failed for {script_name}: {validation_result.get_summary()}")
                    op.details['validation_failed'] = True
                    return False
                elif validation_result.warnings:
                    print(Fore.YELLOW + f"⚠️  Warnings for {script_name}: {'; '.join(validation_result.warnings)}")
            
            # Backup
            backup_info = self.create_script_backup(full_path)
            if backup_info:
                print(Fore.CYAN + f"💾 {backup_info}")
                op.details['backup_created'] = True
            
            # Dry run check
            if is_dry_run():
                print(Fore.YELLOW + f"🔍 DRY RUN: Would execute {rel_path}")
                self.logger.info(f"Dry run: {script_name}")
                op.details['dry_run'] = True
                return True
            
            # User confirmation (unless in testing/minimal mode)
            if not self.env_config.name in ['testing', 'minimal']:
                try:
                    input(Fore.CYAN + f"Ready to run {rel_path}? Press Enter to continue or Ctrl+C to abort...")
                except KeyboardInterrupt:
                    print(Fore.YELLOW + "\nAborted by user")
                    return False
            
            print(Fore.CYAN + f"🚀 Running {rel_path}...")
            
            # Execute script
            start_time = time.time()
            result = subprocess.run(
                ["python", script_name],
                cwd=script_dir,
                capture_output=True,
                text=True,
                shell=False
            )
            duration = time.time() - start_time
            
            # Parse output
            parser = self.script_parsers.get(script_name, self._fallback_parser)
            parser(result.stdout, result.stderr)
            
            # Log results
            op.details.update({
                'exit_code': result.returncode,
                'duration': duration,
                'stdout_lines': len(result.stdout.splitlines()) if result.stdout else 0,
                'stderr_lines': len(result.stderr.splitlines()) if result.stderr else 0
            })
            
            if result.returncode == 0:
                print(Fore.GREEN + f"✅ Completed {rel_path} successfully in {duration:.2f}s")
                print(Fore.CYAN + "-" * 70)
                self.logger.log_operation_complete(f"Script {script_name}", duration, "success")
                self.metrics.increment_counter("successful_scripts")
                return True
            else:
                print(Fore.RED + f"❌ Script {rel_path} failed with exit code {result.returncode}")
                self.logger.error(f"Script {script_name} failed with exit code {result.returncode}")
                self.metrics.increment_counter("failed_scripts")
                return False
    
    def run_all_scripts(self, start_from: int = 0) -> Dict[str, Any]:
        """Run all scripts with comprehensive tracking"""
        try:
            # Find latest monthly folder
            folder = self.get_latest_monthly_folder(".")
            if folder is None:
                return {'success': False, 'error': 'No monthly folder found'}
            
            print(Fore.CYAN + f"📁 Running scripts from: {folder}")
            self.logger.info(f"Starting batch execution from folder: {folder}")
            
            # Gather and sort scripts
            scripts = self.gather_scripts(folder)
            scripts.sort(key=lambda x: self.extract_number(os.path.basename(x[0])))
            
            if not scripts:
                print(Fore.RED + "❌ No Python scripts found")
                return {'success': False, 'error': 'No scripts found'}
            
            # Display available scripts
            print(Fore.CYAN + "\n📋 Available scripts:")
            for idx, (rel_path, _) in enumerate(scripts, start=1):
                print(f"  {idx}: {rel_path}")
            
            # Validate start index
            if not (0 <= start_from < len(scripts)):
                print(Fore.RED + f"Invalid start index {start_from}. Starting from beginning.")
                start_from = 0
            
            # Set up progress tracking
            scripts_to_run = scripts[start_from:]
            if should_show_progress():
                from utils.progress import ProgressTracker
                progress = ProgressTracker(len(scripts_to_run), "Script Execution")
            
            # Execute scripts
            successful = 0
            failed = 0
            
            for i, (rel_path, full_path) in enumerate(scripts_to_run):
                if should_show_progress():
                    progress.update(0, f"Starting {os.path.basename(rel_path)}")
                
                success = self.run_script(rel_path, full_path)
                
                if success:
                    successful += 1
                else:
                    failed += 1
                    # Ask user if they want to continue on failure
                    if not self.env_config.name in ['testing', 'minimal']:
                        try:
                            continue_choice = input(Fore.YELLOW + "Script failed. Continue with next script? (y/N): ")
                            if continue_choice.lower() != 'y':
                                print(Fore.YELLOW + "Stopping execution.")
                                break
                        except KeyboardInterrupt:
                            print(Fore.YELLOW + "\nExecution aborted by user.")
                            break
                
                if should_show_progress():
                    progress.update(1, f"Completed {os.path.basename(rel_path)}")
            
            if should_show_progress():
                progress.complete()
            
            # Complete metrics and show summary
            self.metrics.complete_script()
            
            # Save backup manifest if backups were created
            if self.backup_manager:
                self.backup_manager.save_manifest()
                self.backup_manager.cleanup_old_backups()
            
            # Print final summary
            self.print_execution_summary(successful, failed)
            
            return {
                'success': failed == 0,
                'successful_scripts': successful,
                'failed_scripts': failed,
                'total_scripts': len(scripts_to_run),
                'metrics': self.metrics.get_summary()
            }
            
        except KeyboardInterrupt:
            print(Fore.YELLOW + "\n⚠️  Execution aborted by user.")
            self.logger.warning("Execution aborted by user")
            return {'success': False, 'error': 'Aborted by user'}
        except Exception as e:
            print(Fore.RED + f"❌ Unexpected error: {e}")
            self.logger.error(f"Unexpected error in orchestrator: {e}")
            return {'success': False, 'error': str(e)}
    
    def print_execution_summary(self, successful: int, failed: int):
        """Print execution summary"""
        print("\n" + "="*70)
        print("🎯 EXECUTION SUMMARY")
        print("="*70)
        print(f"✅ Successful Scripts: {successful}")
        print(f"❌ Failed Scripts: {failed}")
        print(f"📊 Success Rate: {(successful/(successful+failed)*100):.1f}%" if (successful+failed) > 0 else "N/A")
        
        if self.metrics:
            summary = self.metrics.get_summary()
            print(f"⏱️  Total Time: {summary['total_time']:.2f}s")
            print(f"🔄 Total Operations: {summary['total_operations']}")
            
            if summary['counters']:
                print("📈 Counters:")
                for name, count in summary['counters'].items():
                    print(f"   {name}: {count}")
        
        if self.backup_manager:
            backup_summary = self.backup_manager.get_backup_summary()
            print(f"💾 Backups Created: {backup_summary['total_backups']}")
        
        print("="*70)

def main():
    """Main entry point"""
    print(Fore.CYAN + "🚀 Enhanced Qualys Automation Orchestrator")
    
    # Show environment info
    print_environment_info()
    
    # Create orchestrator
    orchestrator = EnhancedOrchestrator()
    
    # Get starting script choice
    try:
        choice = input(Fore.CYAN + "\nEnter script number to start from (1 for first script, Enter for 1): ").strip()
        start_idx = int(choice) - 1 if choice else 0
    except ValueError:
        print(Fore.YELLOW + "Invalid input. Starting from beginning.")
        start_idx = 0
    
    # Run scripts
    result = orchestrator.run_all_scripts(start_idx)
    
    # Exit with appropriate code
    exit_code = 0 if result['success'] else 1
    sys.exit(exit_code)

if __name__ == "__main__":
    main() 