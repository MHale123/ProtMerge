"""
excel_formatter.py for ProtMerge v1.2.0
"""

import logging
import time
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from config import OUTPUT_COLUMNS, AMINO_ACID_COLUMNS, PDB_COLUMNS, SIMILARITY_COLUMNS, THEMES


class ExcelFormatter:
    """
    Handles Excel output formatting and saving with professional styling.
    Completely rewritten with proper error handling and similarity support.
    """
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.logger.info("ExcelFormatter initialized")
    
    def save_results(self, input_file, results, options):
        """
        Save results to professionally formatted Excel file with multiple sheets.
        
        Args:
            input_file: Path to original input file
            results: DataFrame with analysis results
            options: Dictionary with analysis options
            
        Returns:
            Path to output file or None if failed
        """
        try:
            self.logger.info(f"Creating Excel output for {len(results)} proteins")
            
            # Prepare output file path
            output_file = self._prepare_output_file(input_file)
            
            # Create workbook
            wb = self._create_workbook(input_file)
            
            # Always create main results sheet
            self._create_main_sheet(wb, results)
            
            # Create optional sheets based on options and data availability
            if options.get('amino_acid', False) and self._has_amino_acid_data(results):
                self._create_amino_acid_sheet(wb, results)
            
            if options.get('pdb_search', False) and self._has_pdb_data(results):
                self._create_pdb_sheet(wb, results)
            
            # Create similarity sheet if results are available
            if options.get('similarity_results') is not None:
                self._create_similarity_sheet(wb, results, options)
            
            # Save workbook
            return self._save_workbook(wb, output_file)
            
        except Exception as e:
            self.logger.error(f"Excel formatting failed: {e}")
            import traceback
            self.logger.error(f"Traceback: {traceback.format_exc()}")
            return self._create_emergency_backup(results, input_file)
    
    def _prepare_output_file(self, input_file):
        """Prepare output file path with conflict resolution"""
        input_path = Path(input_file)
        output_file = input_path.parent / f"{input_path.stem}_ProtMerge.xlsx"
        
        # Handle file conflicts
        counter = 1
        while output_file.exists():
            try:
                # Test if file is accessible
                with open(output_file, 'r+b'):
                    break
            except PermissionError:
                # File is open, try alternative name
                output_file = input_path.parent / f"{input_path.stem}_ProtMerge_{counter}.xlsx"
                counter += 1
                if counter > 10:
                    break
        
        return output_file
    
    def _create_workbook(self, input_file):
        """Create or load workbook"""
        try:
            wb = load_workbook(input_file)
            self.logger.debug("Loaded existing workbook")
        except Exception:
            wb = Workbook()
            # Remove default sheet
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb['Sheet'])
            self.logger.debug("Created new workbook")
        
        return wb
    
    def _create_main_sheet(self, wb, results):
        """Create main results sheet with proper formatting"""
        # Remove existing sheet if it exists
        if 'ProtMerge_Results' in wb.sheetnames:
            wb.remove(wb['ProtMerge_Results'])
        
        ws = wb.create_sheet('ProtMerge_Results')
        
        # Prepare data for output
        output_df = self._prepare_main_data(results)
        
        # Write to sheet
        self._write_dataframe_to_sheet(ws, output_df)
        self._format_sheet(ws, THEMES['main'])
        
        self.logger.info(f"Created main sheet with {len(output_df)} entries")
    
    def _prepare_main_data(self, results):
        """Prepare main results data for Excel output"""
        output_df = pd.DataFrame()
        output_df['UniProt ID'] = results['UniProt_ID']
        
        # Add basic output columns that exist in results
        for internal_key, excel_column in OUTPUT_COLUMNS.items():
            if internal_key in results.columns:
                output_df[excel_column] = results[internal_key]
            else:
                output_df[excel_column] = "NO VALUE FOUND"
        
        return output_df
    
    def _create_amino_acid_sheet(self, wb, results):
        """Create amino acid composition sheet"""
        if 'Amino_Acid_Composition' in wb.sheetnames:
            wb.remove(wb['Amino_Acid_Composition'])
        
        ws = wb.create_sheet('Amino_Acid_Composition')
        
        # Prepare amino acid data
        aa_df = self._prepare_amino_acid_data(results)
        
        # Write to sheet
        self._write_dataframe_to_sheet(ws, aa_df)
        self._format_sheet(ws, THEMES['amino'])
        
        has_data = self._has_amino_acid_data(results)
        self.logger.info(f"Created amino acid sheet ({'with data' if has_data else 'no data available'})")
    
    def _prepare_amino_acid_data(self, results):
        """Prepare amino acid composition data"""
        aa_df = pd.DataFrame()
        aa_df['UniProt ID'] = results['UniProt_ID']
        
        # Add gene name if available
        if 'gene_name' in results.columns:
            aa_df['Gene Name'] = results['gene_name']
        else:
            aa_df['Gene Name'] = "N/A"
        
        # Add amino acid columns
        for internal_key, excel_column in AMINO_ACID_COLUMNS.items():
            if internal_key in results.columns:
                aa_df[excel_column] = results[internal_key]
            else:
                aa_df[excel_column] = "NO VALUE FOUND"
        
        return aa_df
    
    def _create_pdb_sheet(self, wb, results):
        """Create PDB structural analysis sheet"""
        if 'PDB_Structural_Analysis' in wb.sheetnames:
            wb.remove(wb['PDB_Structural_Analysis'])
        
        ws = wb.create_sheet('PDB_Structural_Analysis')
        
        # Prepare PDB data
        pdb_df = self._prepare_pdb_data(results)
        
        # Write to sheet
        self._write_dataframe_to_sheet(ws, pdb_df)
        self._format_sheet(ws, THEMES['pdb'])
        
        has_data = self._has_pdb_data(results)
        self.logger.info(f"Created PDB sheet ({'with data' if has_data else 'no data available'})")
    
    def _prepare_pdb_data(self, results):
        """Prepare PDB structural data"""
        pdb_df = pd.DataFrame()
        pdb_df['UniProt ID'] = results['UniProt_ID']
        
        # Add gene name if available
        if 'gene_name' in results.columns:
            pdb_df['Gene Name'] = results['gene_name']
        else:
            pdb_df['Gene Name'] = "N/A"
        
        # Add PDB columns
        for internal_key, excel_column in PDB_COLUMNS.items():
            if internal_key in results.columns:
                pdb_df[excel_column] = results[internal_key]
            else:
                pdb_df[excel_column] = "NO VALUE FOUND"
        
        return pdb_df
    
    def _create_similarity_sheet(self, wb, results, options):
        """
        Create similarity analysis sheet - FIXED VERSION with proper self parameter
        
        Args:
            wb: Workbook object
            results: Main results DataFrame
            options: Analysis options containing similarity results
        """
        if 'Similarity_Analysis' in wb.sheetnames:
            wb.remove(wb['Similarity_Analysis'])
        
        ws = wb.create_sheet('Similarity_Analysis')
        
        # Get similarity results from options
        similarity_results = options.get('similarity_results', None)
        central_protein = options.get('central_protein_id', 'Unknown')
        
        if similarity_results is not None and not similarity_results.empty:
            # Create comprehensive similarity data
            sim_df = self._prepare_similarity_data(similarity_results, central_protein)
        else:
            # Create placeholder data for failed analysis
            sim_df = self._create_similarity_placeholder(central_protein)
        
        # Write to sheet
        self._write_dataframe_to_sheet(ws, sim_df)
        self._format_sheet(ws, THEMES['similarity'])
        
        # Add conditional formatting for similarity scores
        self._add_similarity_formatting(ws)
        
        has_data = similarity_results is not None and not similarity_results.empty
        self.logger.info(f"Created similarity sheet ({'with data' if has_data else 'no data available'})")
    
    def _prepare_similarity_data(self, similarity_results, central_protein):
        """Prepare similarity results data for Excel"""
        # Create main results data
        sim_df = pd.DataFrame()
        
        # Add ranking and core information
        sim_df['Rank'] = range(1, len(similarity_results) + 1)
        sim_df['Protein ID'] = similarity_results['protein_id']
        sim_df['Overall Similarity'] = similarity_results['overall_similarity'].round(4)
        sim_df['Data Quality'] = similarity_results['data_quality'].round(3)
        
        # Add category breakdowns if available
        category_columns = [
            'sequence_length', 'sequence_identity', 'molecular_weight',
            'isoelectric_point', 'gravy_score', 'functional_keywords',
            'organism_similarity', 'amino_acid_composition'
        ]
        
        for col in category_columns:
            if col in similarity_results.columns:
                display_name = col.replace('_', ' ').title()
                sim_df[display_name] = similarity_results[col].round(3)
        
        # Add metadata header rows
        metadata_rows = pd.DataFrame({
            'Rank': ['Analysis Info:', 'Central Protein:', 'Analysis Date:', 'Total Proteins:', ''],
            'Protein ID': ['Similarity Analysis Results', central_protein, 
                          pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"), 
                          str(len(similarity_results)), 'Results:'],
            'Overall Similarity': ['', '', '', '', ''],
            'Data Quality': ['', '', '', '', '']
        })
        
        # Combine metadata and results
        final_df = pd.concat([metadata_rows, sim_df], ignore_index=True)
        
        return final_df
    
    def _create_similarity_placeholder(self, central_protein):
        """Create placeholder data when similarity analysis fails"""
        return pd.DataFrame({
            'Rank': ['Analysis Failed'],
            'Protein ID': [f'Central Protein: {central_protein}'],
            'Overall Similarity': ['No results available'],
            'Data Quality': ['Check analysis parameters']
        })
    
    def _add_similarity_formatting(self, ws):
        """Add conditional formatting for similarity scores"""
        try:
            from openpyxl.formatting.rule import ColorScaleRule
            
            # Find the Overall Similarity column
            for col_idx, cell in enumerate(ws[1], 1):
                if cell.value == 'Overall Similarity':
                    col_letter = chr(64 + col_idx)  # Convert to letter
                    
                    # Apply color scale from red (low) to green (high)
                    color_scale = ColorScaleRule(
                        start_type='min', start_color='FFFF0000',  # Red
                        mid_type='percentile', mid_value=50, mid_color='FFFFFF00',  # Yellow
                        end_type='max', end_color='FF00FF00'  # Green
                    )
                    
                    ws.conditional_formatting.add(f'{col_letter}6:{col_letter}{ws.max_row}', color_scale)
                    break
                    
        except ImportError:
            # openpyxl conditional formatting not available
            self.logger.warning("Conditional formatting not available")
        except Exception as e:
            self.logger.warning(f"Could not add similarity formatting: {e}")
    
    def _has_amino_acid_data(self, results):
        """Check if amino acid data is available"""
        aa_keys = list(AMINO_ACID_COLUMNS.keys())
        return self._has_data_for_fields(results, aa_keys)
    
    def _has_pdb_data(self, results):
        """Check if PDB data is available"""
        pdb_keys = list(PDB_COLUMNS.keys())
        return self._has_data_for_fields(results, pdb_keys)
    
    def _has_data_for_fields(self, results, field_keys):
        """Check if any of the specified fields have actual data"""
        for key in field_keys:
            if key in results.columns:
                has_data = results[key].apply(
                    lambda x: (pd.notna(x) and 
                              str(x) != "NO VALUE FOUND" and 
                              str(x).strip() != "" and
                              str(x).lower() != "nan")
                ).any()
                if has_data:
                    return True
        return False
    
    def _write_dataframe_to_sheet(self, ws, df):
        """Write DataFrame to worksheet"""
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
    
    def _format_sheet(self, ws, theme):
        """Apply professional formatting to worksheet"""
        if not ws.max_row or ws.max_row < 1:
            return
        
        # Define styles
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color=theme['header'], end_color=theme['header'], fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        data_font = Font(name="Segoe UI", size=10)
        data_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        alt_fill = PatternFill(start_color=theme['alt_row'], end_color=theme['alt_row'], fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin', color='E1E5E9'),
            right=Side(style='thin', color='E1E5E9'),
            top=Side(style='thin', color='E1E5E9'),
            bottom=Side(style='thin', color='E1E5E9')
        )
        
        # Format header row
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        
        # Format data rows
        for row_num in range(2, ws.max_row + 1):
            fill = alt_fill if row_num % 2 == 0 else PatternFill()
            
            for cell in ws[row_num]:
                cell.font = data_font
                cell.alignment = data_alignment
                cell.border = thin_border
                cell.fill = fill
        
        # Auto-size columns
        self._auto_size_columns(ws)
        
        # Freeze header row
        ws.freeze_panes = 'A2'
        
        # Make hyperlinks clickable
        self._make_hyperlinks_clickable(ws)
    
    def _auto_size_columns(self, ws):
        """Auto-size columns with reasonable limits"""
        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            
            for cell in col:
                if cell.value:
                    try:
                        length = len(str(cell.value))
                        max_length = max(max_length, length)
                    except:
                        pass
            
            # Set width with limits
            width = min(max(max_length + 2, 12), 45)
            ws.column_dimensions[column_letter].width = width
    
    def _make_hyperlinks_clickable(self, ws):
        """Make URLs clickable"""
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('https://'):
                    cell.hyperlink = cell.value
                    cell.font = Font(name="Segoe UI", size=10, color="0563C1", underline="single")
    
    def _save_workbook(self, wb, output_file):
        """Save workbook with error handling"""
        try:
            wb.save(output_file)
            sheets = [ws for ws in wb.sheetnames 
                     if any(name in ws for name in ['ProtMerge', 'Amino', 'PDB', 'Similarity'])]
            self.logger.info(f"Results saved to {output_file}")
            self.logger.info(f"Created sheets: {', '.join(sheets)}")
            return output_file
            
        except PermissionError:
            # Try with timestamp
            timestamp = int(time.time())
            timestamped_file = output_file.parent / f"{output_file.stem}_{timestamp}.xlsx"
            try:
                wb.save(timestamped_file)
                self.logger.info(f"Results saved to timestamped file: {timestamped_file}")
                return timestamped_file
            except Exception as e:
                self.logger.error(f"Timestamped save failed: {e}")
                return None
        
        except Exception as e:
            self.logger.error(f"Save failed: {e}")
            return None
    
    def _create_emergency_backup(self, results, input_file):
        """Create emergency backup when main save fails"""
        try:
            # Find writable directory
            backup_dirs = [
                Path.home() / "Desktop",
                Path.home() / "Documents", 
                Path.cwd()
            ]
            
            backup_dir = None
            for directory in backup_dirs:
                if directory.exists():
                    try:
                        # Test write access
                        test_file = directory / "test_write.tmp"
                        test_file.touch()
                        test_file.unlink()
                        backup_dir = directory
                        break
                    except:
                        continue
            
            if backup_dir is None:
                backup_dir = Path.cwd()
            
            backup_file = backup_dir / f"ProtMerge_Emergency_Backup_{int(time.time())}.csv"
            
            # Create simple CSV backup
            backup_df = pd.DataFrame()
            backup_df['UniProt_ID'] = results['UniProt_ID']
            
            # Add all available columns
            for col in results.columns:
                if col != 'UniProt_ID':
                    backup_df[col] = results[col]
            
            backup_df.to_csv(backup_file, index=False)
            self.logger.info(f"Emergency backup saved to: {backup_file}")
            return backup_file
            
        except Exception as e:
            self.logger.error(f"Emergency backup failed: {e}")
            return None
    
    def add_similarity_results_to_options(self, options, similarity_results, central_protein_id):
        """
        Add similarity results to options for Excel export.
        
        Args:
            options: Existing options dictionary
            similarity_results: DataFrame with similarity results
            central_protein_id: ID of central protein
            
        Returns:
            Updated options dictionary
        """
        options = options.copy()
        options['similarity_results'] = similarity_results
        options['central_protein_id'] = central_protein_id
        options['similarity_analysis'] = True
        
        return options


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def format_similarity_results_for_excel(similarity_results, central_protein):
    """
    Format similarity results for Excel export.
    
    Args:
        similarity_results: DataFrame with similarity results
        central_protein: ID of central protein
        
    Returns:
        Formatted DataFrame ready for Excel export
    """
    if similarity_results.empty:
        return pd.DataFrame({
            'Status': ['No Results'],
            'Central Protein': [central_protein],
            'Message': ['Similarity analysis produced no results']
        })
    
    # Format main results
    formatted_df = pd.DataFrame()
    formatted_df['Rank'] = range(1, len(similarity_results) + 1)
    formatted_df['Protein ID'] = similarity_results['protein_id']
    formatted_df['Overall Similarity'] = similarity_results['overall_similarity'].round(4)
    formatted_df['Data Quality Score'] = similarity_results['data_quality'].round(3)
    
    # Add category scores if available
    category_columns = [
        'sequence_length', 'molecular_weight', 'isoelectric_point',
        'gravy_score', 'functional_keywords', 'organism_similarity'
    ]
    
    for col in category_columns:
        if col in similarity_results.columns:
            display_name = col.replace('_', ' ').title()
            formatted_df[display_name] = similarity_results[col].round(3)
    
    return formatted_df


def create_similarity_summary(similarity_results, central_protein):
    """
    Create summary statistics for similarity analysis.
    
    Args:
        similarity_results: DataFrame with similarity results
        central_protein: ID of central protein
        
    Returns:
        Dictionary with summary statistics
    """
    if similarity_results.empty:
        return {
            'central_protein': central_protein,
            'total_proteins': 0,
            'mean_similarity': 0.0,
            'max_similarity': 0.0,
            'min_similarity': 0.0,
            'high_similarity_count': 0
        }
    
    similarities = similarity_results['overall_similarity']
    
    return {
        'central_protein': central_protein,
        'total_proteins': len(similarity_results),
        'mean_similarity': similarities.mean(),
        'max_similarity': similarities.max(),
        'min_similarity': similarities.min(),
        'high_similarity_count': (similarities > 0.7).sum(),
        'analysis_timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
    }