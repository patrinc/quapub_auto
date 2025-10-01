import pandas as pd
import os
import warnings
import sys
from pathlib import Path

# Add utils to path
utils_path = Path(__file__).parent.parent / 'utils'
if str(utils_path) not in sys.path:
    sys.path.insert(0, str(utils_path))

# Import enhanced utilities
from utils import (
    setup_logger, ProcessingMetrics, safe_read_excel, 
    normalize_key_columns, filter_dataframe_by_criteria,
    ensure_directory_exists, BackupManager, get_config
)

# Import enhanced config
try:
    sys.path.insert(0, str(Path(__file__).parent.parent))
    import enhanced_config as config
    print("✅ Using enhanced configuration")
except ImportError:
    import config
    print("⚠️  Using basic configuration")

warnings.simplefilter(action='ignore', category=FutureWarning)

class ReportProcessor:
    """Enhanced report processor with comprehensive logging and metrics"""
    
    def __init__(self):
        self.env_config = get_config()
        self.logger = setup_logger('process_reports', self.env_config.log_level)
        self.metrics = ProcessingMetrics('process_reports')
        self.backup_manager = BackupManager(logger=self.logger) if self.env_config.create_backups else None
        
        # Configuration
        self.countries = getattr(config, 'COUNTRIES', ['Brazil', 'CSA', 'India', 'Mexico', 'SFTL'])
        self.report_types = getattr(config, 'REPORT_TYPES', ['Network', 'Server'])
        self.key_cols = getattr(config, 'KEY_COLUMNS', ['IP', 'QID', 'Port'])
        
        # Processing criteria
        self.filter_criteria = getattr(config, 'FILTER_CRITERIA', {
            'remove_empty_solution': True,
            'remove_types': ['Ig', 'Practice'],
            'remove_severity': [1],
            'default_status': 'Reviewed'
        })
        
        self.logger.info("Report processor initialized")
        self.logger.info(f"Countries: {', '.join(self.countries)}")
        self.logger.info(f"Report types: {', '.join(self.report_types)}")
    
    def get_file_paths(self, country: str, report_type: str):
        """Get input and output file paths"""
        if hasattr(config, 'build_raw_report_filename'):
            input_filename = config.build_raw_report_filename(country, report_type)
            input_path = config.get_monthly_path() / input_filename
        else:
            # Fallback to original logic
            year_short = config.year[-2:]
            curr_raw_folder = f"{year_short}_{config.current_month_num} Monthly Scans"
            base_path = getattr(config, 'base_path', r"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning")
            input_path = Path(base_path) / curr_raw_folder / f"{country}_{config.year}_{config.current_month_num}_Raw_Report_{report_type}_Report.xlsx"
        
        # Output path
        if hasattr(config, 'get_status_overview_path'):
            output_dir = config.get_status_overview_path()
        else:
            output_dir = input_path.parent / "01_with_Status_and_Overview"
        
        ensure_directory_exists(output_dir)
        
        if hasattr(config, 'build_final_report_filename'):
            output_filename = config.build_final_report_filename(country, report_type, with_status=True)
        else:
            output_filename = f"{country}_{config.year}_{config.current_month_num}_Final_Report_with_Status_{report_type}_Report.xlsx"
        
        output_path = output_dir / output_filename
        
        return input_path, output_path
    
    def process_single_report(self, country: str, report_type: str) -> bool:
        """Process a single country-report type combination"""
        operation_name = f"process_{country}_{report_type}"
        
        with self.metrics.track_operation(operation_name) as op:
            self.logger.log_operation_start(operation_name, f"Processing {country} {report_type} report")
            
            try:
                # Get file paths
                input_path, output_path = self.get_file_paths(country, report_type)
                op.details.update({
                    'input_path': str(input_path),
                    'output_path': str(output_path),
                    'country': country,
                    'report_type': report_type
                })
                
                # Load data
                curr_raw = safe_read_excel(input_path, self.logger)
                
                if curr_raw.empty:
                    self.logger.warning(f"Skipping {country} - {report_type} due to missing input file")
                    print(f"[{country} - {report_type}] Skipping due to missing input file.\n")
                    op.details['skipped'] = True
                    return True  # Not a failure, just no data
                
                original_count = len(curr_raw)
                op.details['original_rows'] = original_count
                self.metrics.increment_counter('total_rows_processed', original_count)
                
                # Normalize key columns
                curr_raw = normalize_key_columns(curr_raw, self.key_cols, self.logger)
                
                # Apply filtering
                filtered_df = self.apply_filters(curr_raw, country, report_type)
                op.details['filtered_rows'] = len(filtered_df)
                
                # Add status column
                filtered_df['Status'] = self.filter_criteria.get('default_status', 'Reviewed')
                
                # Ensure Comments column exists
                if 'Comments' not in filtered_df.columns:
                    insert_pos = filtered_df.columns.get_loc('Status') + 1
                    filtered_df.insert(insert_pos, 'Comments', '')
                
                # Sort by key columns for consistency
                final_sorted = filtered_df.sort_values(by=self.key_cols)
                
                # Create backup if enabled
                if self.backup_manager and output_path.exists():
                    backup_path = self.backup_manager.backup_file(output_path, 'processing')
                    self.logger.info(f"Created backup: {backup_path}")
                
                # Save final report
                final_sorted.to_excel(output_path, index=False)
                
                # Log success
                final_count = len(final_sorted)
                removed_count = original_count - final_count
                
                self.logger.log_file_processed(str(output_path), final_count)
                self.logger.info(f"Processed {country} {report_type}: {original_count} -> {final_count} rows ({removed_count} removed)")
                
                print(f"[{country} - {report_type}] Final report saved: {output_path.name}")
                print(f"[{country} - {report_type}] Rows: {original_count} -> {final_count} (removed {removed_count})\n")
                
                op.details.update({
                    'final_rows': final_count,
                    'removed_rows': removed_count,
                    'success': True
                })
                
                self.metrics.increment_counter('reports_processed')
                self.metrics.increment_counter('rows_removed', removed_count)
                
                return True
                
            except Exception as e:
                self.logger.log_error_with_context(e, f"processing {country} {report_type}")
                print(f"[{country} - {report_type}] Error: {e}\n")
                op.details.update({
                    'error': str(e),
                    'success': False
                })
                return False
    
    def apply_filters(self, df: pd.DataFrame, country: str, report_type: str) -> pd.DataFrame:
        """Apply filtering criteria to dataframe"""
        with self.metrics.track_operation(f"filter_{country}_{report_type}"):
            original_count = len(df)
            filtered_df = df.copy()
            
            # Filter out rows with empty Solution
            if self.filter_criteria.get('remove_empty_solution'):
                before_filter = len(filtered_df)
                filtered_df = filtered_df[~(filtered_df['Solution'].isna() | 
                                          (filtered_df['Solution'].astype(str).str.strip() == ''))]
                removed_empty_solution = before_filter - len(filtered_df)
                
                if removed_empty_solution > 0:
                    print(f"[{country} - {report_type}] Removed {removed_empty_solution} rows with empty 'Solution'.")
                    self.logger.info(f"Removed {removed_empty_solution} rows with empty 'Solution'")
            
            # Remove rows with specific types
            remove_types = self.filter_criteria.get('remove_types', [])
            if remove_types and 'Type' in filtered_df.columns:
                before_type_filter = len(filtered_df)
                filtered_df = filtered_df[~filtered_df['Type'].isin(remove_types)]
                removed_type = before_type_filter - len(filtered_df)
                
                if removed_type > 0:
                    print(f"[{country} - {report_type}] Removed {removed_type} rows with Type in {remove_types}.")
                    self.logger.info(f"Removed {removed_type} rows with Type in {remove_types}")
            
            # Remove rows with specific severity levels
            remove_severity = self.filter_criteria.get('remove_severity', [])
            if remove_severity and 'Severity' in filtered_df.columns:
                before_severity_filter = len(filtered_df)
                filtered_df = filtered_df[~filtered_df['Severity'].isin(remove_severity)]
                removed_severity = before_severity_filter - len(filtered_df)
                
                if removed_severity > 0:
                    print(f"[{country} - {report_type}] Removed {removed_severity} rows with Severity in {remove_severity}.")
                    self.logger.info(f"Removed {removed_severity} rows with Severity in {remove_severity}")
            
            total_removed = original_count - len(filtered_df)
            if total_removed > 0:
                self.logger.info(f"Total filtering removed {total_removed} rows for {country} {report_type}")
            
            return filtered_df
    
    def process_all_reports(self) -> bool:
        """Process all country-report type combinations"""
        self.logger.info("Starting batch processing of all reports")
        print("Starting enhanced report processing...")
        
        total_combinations = len(self.countries) * len(self.report_types)
        successful = 0
        failed = 0
        
        # Set up progress tracking if enabled
        if self.env_config.enable_progress:
            from utils.progress import ProgressTracker
            progress = ProgressTracker(total_combinations, "Processing Reports")
        
        for country in self.countries:
            for report_type in self.report_types:
                if self.env_config.enable_progress:
                    progress.update(0, f"{country} {report_type}")
                
                print(f"Processing {country} - {report_type} reports...")
                success = self.process_single_report(country, report_type)
                
                if success:
                    successful += 1
                else:
                    failed += 1
                
                if self.env_config.enable_progress:
                    progress.update(1, f"Completed {country} {report_type}")
        
        if self.env_config.enable_progress:
            progress.complete()
        
        # Complete processing and show summary
        self.metrics.complete_script()
        
        # Save backup manifest if backups were created
        if self.backup_manager:
            self.backup_manager.save_manifest()
        
        # Print summary
        self.print_summary(successful, failed)
        
        return failed == 0
    
    def print_summary(self, successful: int, failed: int):
        """Print processing summary"""
        print("\n" + "="*60)
        print("📊 REPORT PROCESSING SUMMARY")
        print("="*60)
        print(f"✅ Successful: {successful}")
        print(f"❌ Failed: {failed}")
        print(f"📈 Success Rate: {(successful/(successful+failed)*100):.1f}%" if (successful+failed) > 0 else "N/A")
        
        summary = self.metrics.get_summary()
        print(f"⏱️  Total Time: {summary['total_time']:.2f}s")
        
        if summary['counters']:
            print("📊 Processing Stats:")
            for name, count in summary['counters'].items():
                print(f"   {name}: {count:,}")
        
        print("="*60)
        
        self.logger.info(f"Processing completed: {successful} successful, {failed} failed")

def main():
    """Main entry point"""
    processor = ReportProcessor()
    
    try:
        success = processor.process_all_reports()
        if success:
            print("All reports processed successfully.")
        else:
            print("Some reports failed to process.")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\nProcessing interrupted by user.")
        processor.logger.warning("Processing interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        processor.logger.error(f"Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 