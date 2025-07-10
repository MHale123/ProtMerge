#!/usr/bin/env python3
"""
ProtMerge - Protein Analysis Tool
Version: 1.2.0
Created by Matthew Hale
"""

import sys
import logging
import time
from pathlib import Path
from gui_main import ProtMergeGUI
from data_handler import DataHandler
from analyzers import AnalyzerManager
from excel_formatter import ExcelFormatter

class ProtMerge:
    """Main ProtMerge application class"""
    
    def __init__(self):
        self.data_handler = DataHandler()
        self.analyzer_manager = AnalyzerManager()
        self.excel_formatter = ExcelFormatter()
        self.logger = logging.getLogger(__name__)
        self.analysis_summary = None
        self.logger.info("ProtMerge v1.2.0 initialized")
        
    def run_gui(self):
        """Launch the GUI interface"""
        gui = ProtMergeGUI(self)
        gui.run()
    
    def run_analysis(self, input_file, sheet_name, column_index, options, progress_callback=None):
        """Run complete analysis pipeline"""
        try:
            self.logger.info(f"Starting analysis for {Path(input_file).name}")
            
            # Load data
            if progress_callback:
                progress_callback(0, "Loading data", f"Reading {Path(input_file).name}")
            
            data = self.data_handler.load_excel_data(
                input_file, sheet_name, column_index, options.get('safe_mode', True)
            )
            
            protein_count = len(data['results'])
            self.logger.info(f"Loaded {protein_count} proteins for analysis")
            
            # Run analyses
            if progress_callback:
                progress_callback(5, "Running analyses", "Starting protein data collection")
            
            results = self.analyzer_manager.run_all_analyses(data, options, progress_callback)
            
            # Calculate analysis summary
            self.analysis_summary = self._calculate_analysis_summary(results, options)
            
            # Save results
            if progress_callback:
                progress_callback(98, "Saving results", "Creating Excel file")
            
            output_file = self.excel_formatter.save_results(input_file, results, options)
            
            # Log completion summary
            self._log_completion_summary(output_file, results, options)
            
            if progress_callback:
                progress_callback(100, "Analysis complete", "Results saved successfully")
            
            return output_file
            
        except Exception as e:
            self.logger.error(f"Analysis pipeline failed: {e}")
            raise
    
    def get_analysis_summary(self):
        """Get analysis summary for completion dialog"""
        return self.analysis_summary or {}
    
    def _calculate_analysis_summary(self, results, options):
        """Calculate analysis summary statistics"""
        try:
            total_proteins = len(results)
            
            # Count data completeness
            uniprot_complete = sum(1 for _, row in results.iterrows() 
                                 if self._is_data_complete(row.get('function', '')))
            
            protparam_complete = sum(1 for _, row in results.iterrows() 
                                   if self._is_data_complete(row.get('mw', '')))
            
            blast_complete = sum(1 for _, row in results.iterrows() 
                               if self._is_data_complete(row.get('identity', '')))
            
            pdb_complete = sum(1 for _, row in results.iterrows() 
                             if self._is_data_complete(row.get('structure_count', ''), check_zero=True))
            
            summary = {
                'total_proteins': total_proteins,
                'uniprot_complete': uniprot_complete,
                'uniprot_percent': (uniprot_complete / total_proteins) * 100 if total_proteins > 0 else 0,
                'protparam_complete': protparam_complete if options.get('protparam', False) else 0,
                'protparam_percent': (protparam_complete / total_proteins) * 100 if total_proteins > 0 else 0,
                'blast_complete': blast_complete if options.get('blast', False) else 0,
                'blast_percent': (blast_complete / total_proteins) * 100 if total_proteins > 0 else 0,
                'pdb_complete': pdb_complete if options.get('pdb_search', False) else 0,
                'pdb_percent': (pdb_complete / total_proteins) * 100 if total_proteins > 0 else 0
            }
            
            return summary
            
        except Exception as e:
            self.logger.error(f"Error calculating analysis summary: {e}")
            return {
                'total_proteins': len(results) if results is not None else 0,
                'uniprot_complete': 0,
                'uniprot_percent': 0,
                'protparam_complete': 0,
                'protparam_percent': 0,
                'blast_complete': 0,
                'blast_percent': 0,
                'pdb_complete': 0,
                'pdb_percent': 0
            }
    
    def _is_data_complete(self, value, check_zero=False):
        """Check if data value indicates completion"""
        if value is None:
            return False
        
        value_str = str(value).strip().upper()
        
        # Check for no value indicators
        if value_str in ['', 'NO VALUE FOUND', 'NAN', 'NONE', 'N/A']:
            return False
        
        # For PDB structure count, check if it's a positive number
        if check_zero:
            try:
                num_value = float(value_str)
                return num_value > 0
            except (ValueError, TypeError):
                return False
        
        return True
    
    def _log_completion_summary(self, output_file, results, options):
        """Log analysis completion summary"""
        try:
            summary = self.analysis_summary
            
            self.logger.info("=" * 60)
            self.logger.info("ANALYSIS COMPLETION SUMMARY")
            self.logger.info("=" * 60)
            self.logger.info(f"Total proteins analyzed: {summary['total_proteins']}")
            self.logger.info(f"UniProt data: {summary['uniprot_complete']}/{summary['total_proteins']} ({summary['uniprot_percent']:.1f}%)")
            
            if options.get('protparam', False):
                self.logger.info(f"ProtParam data: {summary['protparam_complete']}/{summary['total_proteins']} ({summary['protparam_percent']:.1f}%)")
            
            if options.get('blast', False):
                self.logger.info(f"BLAST data: {summary['blast_complete']}/{summary['total_proteins']} ({summary['blast_percent']:.1f}%)")
            
            if options.get('pdb_search', False):
                self.logger.info(f"PDB structures: {summary['pdb_complete']}/{summary['total_proteins']} ({summary['pdb_percent']:.1f}%)")
            
            if output_file:
                self.logger.info(f"Results saved to: {output_file}")
                
                # List created sheets
                try:
                    from openpyxl import load_workbook
                    wb = load_workbook(output_file)
                    sheets = [ws for ws in wb.sheetnames if any(name in ws for name in ['ProtMerge', 'Amino', 'PDB'])]
                    self.logger.info(f"Excel sheets created: {', '.join(sheets)}")
                except Exception:
                    pass
            
            self.logger.info("=" * 60)
            
        except Exception as e:
            self.logger.error(f"Error generating completion summary: {e}")
    
    def check_dependencies(self):
        """Check if all dependencies are available"""
        missing_deps = []
        
        # Check core dependencies
        core_deps = ['pandas', 'requests', 'openpyxl', 'lxml']
        for dep in core_deps:
            try:
                __import__(dep)
            except ImportError:
                missing_deps.append(dep)
        
        if missing_deps:
            self.logger.error(f"Missing core dependencies: {', '.join(missing_deps)}")
            return False
        
        self.logger.info("All core dependencies available")
        return True

def setup_logging():
    """Setup logging configuration"""
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    # Create logs directory
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        handlers=[
            logging.FileHandler(log_dir / f"protmerge_{int(time.time())}.log"),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    # Reduce noise from external libraries
    logging.getLogger('urllib3').setLevel(logging.WARNING)
    logging.getLogger('requests').setLevel(logging.WARNING)

def main():
    """Main entry point"""
    try:        
        # Setup logging
        setup_logging()
        logger = logging.getLogger(__name__)
        
        logger.info("Starting ProtMerge v1.2.0 - Protein Analysis Tool")
        
        # Initialize application
        app = ProtMerge()
        
        # Check dependencies
        if not app.check_dependencies():
            logger.error("Missing required dependencies. Please install them and try again.")
            sys.exit(1)
        
        # Launch GUI
        logger.info("Launching graphical interface...")
        app.run_gui()
        
        logger.info("ProtMerge session ended")
        
    except KeyboardInterrupt:
        print("\nProtMerge interrupted by user")
        sys.exit(0)
    except Exception as e:
        print(f"ProtMerge failed to start: {e}")
        logging.error(f"Fatal error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()