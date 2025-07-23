"""
Configuration and constants for ProtMerge
Version: 1.2.0
"""

# Version and application info
VERSION = "1.2.0"
APP_NAME = "ProtMerge"
AUTHOR = "Matthew Hale"
DESCRIPTION = "Automated Protein Data Integration Tool with Advanced Similarity Analysis"

# API endpoints
UNIPROT_BASE_URL = "https://rest.uniprot.org/uniprotkb"
PROTPARAM_URL = "https://web.expasy.org/cgi-bin/protparam/protparam"
BLAST_URL = "https://blast.ncbi.nlm.nih.gov/Blast.cgi"
PDB_SEARCH_URL = "https://search.rcsb.org/rcsbsearch/v2/query"
PDB_DATA_URL = "https://data.rcsb.org/rest/v1/core"

# Rate limiting (seconds between requests)
RATE_LIMITS = {
    'uniprot': 0.1,
    'protparam': 2.0,
    'blast': 10.0,
    'pdb': 0.5,
    'similarity': 0.1,
    'gene_converter': 0.2 
}

# Default analysis options
DEFAULT_OPTIONS = {
    'uniprot': True,
    'protparam': True,
    'blast': False,
    'amino_acid': False,
    'pdb_search': False,
    'similarity_analysis': False,
    'safe_mode': True,
    'use_gene_ids': False 
}

# Output column mappings - maps internal field names to Excel column headers
OUTPUT_COLUMNS = {
    'original_gene_id': 'Original Gene ID',  
    'organism': 'Organism',
    'gene_name': 'Gene Name', 
    'function': 'Protein Function/Notes',
    'environment': 'Environment Source',
    'sequence': 'Protein Sequence',
    'similar': 'BLAST Similar Proteins',
    'identity': '% Identity (Top Hit)',
    'evalue': 'E-value (Top Hit)',
    'align': 'Alignment Length (Top Hit)',
    'mw': 'ProtParam: MW',
    'pi': 'ProtParam: pI',
    'gravy': 'ProtParam: GRAVY',
    'ext': 'Extinction Coefficient (M-1 cm-1)',
    'alphafold': 'AlphaFold Link',
    'keywords': 'Relevant Keywords',
    'structure': 'Structure Type'
}

# PDB-specific columns
PDB_COLUMNS = {
    'structure_count': 'PDB Structures Available',
    'best_resolution': 'Best Resolution',
    'structure_methods': 'Experimental Methods', 
    'complex_info': 'Complex Information',
    'pdb_ids': 'PDB IDs',
    'best_structure': 'Highest Quality Structure',
    'ligand_info': 'Bound Ligands',
    'structure_quality': 'Structure Quality'
}

# Similarity analysis columns - Advanced v1.2.0 feature
SIMILARITY_COLUMNS = {
    'composite_score': 'Composite Similarity Score',
    'z_score': 'Z-Score (Distance from Central)',
    'similarity_category': 'Similarity Category',
    'data_quality_score': 'Data Quality Score',
    'available_dimensions': 'Available Analysis Dimensions',
    'sequence_similarity': 'Sequence Similarity',
    'biochemical_similarity': 'Biochemical Similarity', 
    'functional_similarity': 'Functional Similarity',
    'amino_acid_similarity': 'Amino Acid Similarity'
}

# Amino acid composition columns
AMINO_ACID_COLUMNS = {
    'ala': 'Ala (A)', 'arg': 'Arg (R)', 'asn': 'Asn (N)', 'asp': 'Asp (D)', 'cys': 'Cys (C)',
    'gln': 'Gln (Q)', 'glu': 'Glu (E)', 'gly': 'Gly (G)', 'his': 'His (H)', 'ile': 'Ile (I)',
    'leu': 'Leu (L)', 'lys': 'Lys (K)', 'met': 'Met (M)', 'phe': 'Phe (F)', 'pro': 'Pro (P)',
    'ser': 'Ser (S)', 'thr': 'Thr (T)', 'trp': 'Trp (W)', 'tyr': 'Tyr (Y)', 'val': 'Val (V)',
    'pyl': 'Pyl (O)', 'sec': 'Sec (U)', 'asx': 'Asx (B)', 'glx': 'Glx (Z)', 'xaa': 'Xaa (X)',
    'atomic_comp': 'Atomic Composition'
}

# UI styling themes
THEMES = {
    'main': {
        'header': '2F5597',      # Dark blue header
        'alt_row': 'F8F9FA',     # Light gray alternating rows
        'text_primary': '212529', # Dark text
        'text_secondary': '6C757D' # Gray text
    },
    'amino': {
        'header': '0F7B0F',      # Green header for amino acid sheet
        'alt_row': 'F0F8F0',     # Light green alternating rows
        'text_primary': '212529',
        'text_secondary': '6C757D'
    },
    'similarity': {
        'header': '20C997',      # Teal header for similarity sheet
        'alt_row': 'E8F5F3',     # Light teal alternating rows
        'text_primary': '212529',
        'text_secondary': '6C757D'
    },
    'pdb': {
        'header': '6F42C1',      # Purple header for PDB sheet
        'alt_row': 'F8F5FF',     # Light purple alternating rows
        'text_primary': '212529',
        'text_secondary': '6C757D'
    },
    'colors': {
        'success': '#0F7B0F',    # Green for success messages
        'warning': '#FFC107',    # Yellow for warnings
        'error': '#DC3545',      # Red for errors
        'info': '#2F5597',       # Blue for info messages
        'disabled': '#6C757D',   # Gray for disabled elements
        'similarity': '#20C997'  # Teal for similarity features
    },
    # GUI theme for similarity analysis - using hex color codes directly
    'similarity_live': {
        'header': '#00bcd4',        # Cyan header
        'alt_row': '#E8F5F3',       # Light teal alternating rows
        'text_primary': '#212529',   # Dark text
        'text_secondary': '#6C757D', # Gray text
        'slider_track': '#3c3c3c',   # Dark gray slider track
        'slider_fill': '#00bcd4',    # Cyan slider fill
        'slider_hover': '#00acc1',   # Hover cyan
        'highlight': '#4caf50',      # Green highlight
        'network_node': '#9c27b0',   # Purple nodes
        'network_edge': '#b0b0b0'    # Gray edges
    }
}

# Body location mapping for environment extraction
BODY_LOCATIONS = {
    'gut': [
        'gut', 'intestinal', 'intestine', 'gastrointestinal', 'enteric', 
        'colon', 'colonic', 'fecal', 'faecal', 'microbiome', 'microbiota', 
        'digestive', 'bowel', 'stomach', 'gastric'
    ],
    'blood': [
        'blood', 'plasma', 'serum', 'vascular', 'circulatory', 'hematopoietic'
    ],
    'brain': [
        'brain', 'neural', 'neuronal', 'cerebral', 'nervous system'
    ],
    'liver': ['liver', 'hepatic', 'hepatocyte'],
    'kidney': ['kidney', 'renal', 'nephron'],
    'lung': ['lung', 'pulmonary', 'respiratory'],
    'skin': ['skin', 'dermal', 'epidermal'],
    'muscle': ['muscle', 'muscular', 'myocyte'],
    'bone': ['bone', 'skeletal', 'osseous'],
    'immune': ['immune', 'immunological', 'lymphoid', 'spleen'],
    'oral': ['oral', 'mouth', 'dental', 'saliva']
}

# Gut bacteria for environment detection
GUT_BACTERIA = [
    'bacteroides', 'lactobacillus', 'bifidobacterium', 
    'clostridium', 'escherichia'
]

# ProtParam regex patterns for data extraction
PROTPARAM_PATTERNS = {
    'mw': [
        r'<strong>Molecular weight:</strong>\s*([0-9,]+\.?[0-9]*)', 
        r'Molecular weight:\s*([0-9,]+\.?[0-9]*)'
    ],
    'pi': [
        r'<strong>Theoretical pI:</strong>\s*([0-9]+\.?[0-9]*)', 
        r'Theoretical pI:\s*([0-9]+\.?[0-9]*)'
    ],
    'gravy': [
        r'<strong>Grand average of hydropathicity \(GRAVY\):</strong>\s*([-]?[0-9]+\.?[0-9]*)', 
        r'hydropathicity.*?:\s*([-]?[0-9]+\.?[0-9]*)', 
        r'GRAVY.*?:\s*([-]?[0-9]+\.?[0-9]*)'
    ],
    'ext': [
        r'Ext\.\s*coefficient\s*([0-9,]+)', 
        r'Extinction\s*coefficient.*?([0-9,]+)', 
        r'coefficient.*?([0-9,]+)'
    ]
}

# Amino acid regex patterns for ProtParam parsing
AMINO_ACID_PATTERNS = {
    'ala': r'Ala \(A\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'arg': r'Arg \(R\)\s+(\d+)\s+(\d+\.?\d*)%',
    'asn': r'Asn \(N\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'asp': r'Asp \(D\)\s+(\d+)\s+(\d+\.?\d*)%',
    'cys': r'Cys \(C\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'gln': r'Gln \(Q\)\s+(\d+)\s+(\d+\.?\d*)%',
    'glu': r'Glu \(E\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'gly': r'Gly \(G\)\s+(\d+)\s+(\d+\.?\d*)%',
    'his': r'His \(H\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'ile': r'Ile \(I\)\s+(\d+)\s+(\d+\.?\d*)%',
    'leu': r'Leu \(L\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'lys': r'Lys \(K\)\s+(\d+)\s+(\d+\.?\d*)%',
    'met': r'Met \(M\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'phe': r'Phe \(F\)\s+(\d+)\s+(\d+\.?\d*)%',
    'pro': r'Pro \(P\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'ser': r'Ser \(S\)\s+(\d+)\s+(\d+\.?\d*)%',
    'thr': r'Thr \(T\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'trp': r'Trp \(W\)\s+(\d+)\s+(\d+\.?\d*)%',
    'tyr': r'Tyr \(Y\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'val': r'Val \(V\)\s+(\d+)\s+(\d+\.?\d*)%',
    'pyl': r'Pyl \(O\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'sec': r'Sec \(U\)\s+(\d+)\s+(\d+\.?\d*)%',
    'asx': r'\(B\)\s+(\d+)\s+(\d+\.?\d*)%', 
    'glx': r'\(Z\)\s+(\d+)\s+(\d+\.?\d*)%',
    'xaa': r'\(X\)\s+(\d+)\s+(\d+\.?\d*)%'
}

# Atomic composition pattern
ATOMIC_FORMULA_PATTERN = r'<strong>Formula:</strong>\s*C<sub>(\d+)</sub>H<sub>(\d+)</sub>N<sub>(\d+)</sub>O<sub>(\d+)</sub>S<sub>(\d+)</sub>'

# Application settings
SETTINGS = {
    'window_size': '750x750',    # Updated for similarity analysis
    'window_resizable': False,
    'progress_update_interval': 0.1,  # seconds
    'max_backup_files': 10,
    'log_level': 'INFO',
    'timeout_seconds': 30,
    'min_sequence_length': 20,
    'blast_max_wait': 300,  # 5 minutes
    'similarity_max_proteins': 500,  # Maximum proteins for similarity analysis
    'visualization_timeout': 120  # 2 minutes for visualization generation
}

# File extensions and types
SUPPORTED_EXCEL_FORMATS = [
    ("Excel files", "*.xlsx *.xls"),
    ("All files", "*.*")
]

# COMPLETE ERROR MESSAGES - FIXED VERSION
ERROR_MESSAGES = {
    'file_not_found': "Could not find the specified file",
    'invalid_excel': "File is not a valid Excel format",
    'no_columns': "No columns found in the selected sheet",
    'no_uniprot_column': "Please select a column containing UniProt IDs",
    'network_error': "Network connection error - please check your internet",
    'analysis_failed': "Analysis failed due to an unexpected error",
    'save_failed': "Could not save results file",
    'similarity_failed': "Similarity analysis encountered an error but main analysis completed",
    'central_protein_not_found': "Central protein ID not found in the dataset",
    'insufficient_data': "Insufficient data for similarity analysis",
    
    # MISSING KEYS - NOW ADDED:
    'FILE_LOAD_ERROR': "Failed to load file",
    'INVALID_PROTMERGE_FILE': "This doesn't appear to be a ProtMerge results file",
    'SHEET_NOT_FOUND': "Expected sheet 'ProtMerge_Results' not found",
    'REQUIRED_COLUMN_MISSING': "Required column 'UniProt ID' not found in results file",
    'INSUFFICIENT_DATA': "Need at least 2 proteins for similarity analysis",
    'MODULE_MISSING': "Required module not found",
    'ANALYSIS_FAILED': "Analysis failed",
    'FILE_OPEN_ERROR': "Could not open file:",
    'SIMILARITY_ANALYSIS_FAILED': "Similarity analysis failed"
}

# Success messages
SUCCESS_MESSAGES = {
    'file_loaded': "File loaded successfully",
    'analysis_complete': "Analysis completed successfully",
    'results_saved': "Results saved successfully",
    'similarity_complete': "Similarity analysis completed with interactive visualizations",
    'visualization_created': "Interactive HTML report generated"
}

# Similarity analysis configuration
SIMILARITY_CONFIG = {
    'min_proteins_required': 3,  # Minimum proteins needed for meaningful analysis
    'min_dimensions_required': 2,  # Minimum similarity dimensions for composite score
    'default_weights': {
        'sequence': 0.35,
        'biochemical': 0.25,
        'functional': 0.25,
        'amino_acid': 0.15
    },
    'zscore_categories': {
        'highly_similar': 2.0,
        'similar': 1.0,
        'moderate': -1.0,
        'distant': -2.0
    },
    'quality_thresholds': {
        'excellent': 0.8,
        'good': 0.6,
        'moderate': 0.4,
        'poor': 0.2
    }
}

# Enhanced similarity analysis configuration
SIMILARITY_CATEGORIES = {
    'sequence_properties': {
        'sequence_length': {'weight': 0.10, 'description': 'Sequence length similarity', 'enabled': True},
        'sequence_identity': {'weight': 0.20, 'description': 'Direct sequence similarity', 'enabled': True},
        'alignment_quality': {'weight': 0.10, 'description': 'BLAST alignment quality', 'enabled': True},
    },
    'physicochemical': {
        'molecular_weight': {'weight': 0.08, 'description': 'Molecular weight similarity', 'enabled': True},
        'isoelectric_point': {'weight': 0.06, 'description': 'Isoelectric point similarity', 'enabled': True},
        'gravy_score': {'weight': 0.04, 'description': 'Hydrophobicity similarity', 'enabled': True},
        'extinction_coefficient': {'weight': 0.02, 'description': 'Extinction coefficient similarity', 'enabled': True},
    },
    'amino_acid_composition': {
        'full_aa_composition': {'weight': 0.15, 'description': 'Complete AA composition profile', 'enabled': True},
        'hydrophobic_content': {'weight': 0.05, 'description': 'Hydrophobic residue content', 'enabled': True},
        'charged_residues': {'weight': 0.03, 'description': 'Charged residue content', 'enabled': True},
        'aromatic_content': {'weight': 0.02, 'description': 'Aromatic residue content', 'enabled': True},
        'polar_content': {'weight': 0.02, 'description': 'Polar residue content', 'enabled': True},
    },
    'structural_properties': {
        'structure_availability': {'weight': 0.03, 'description': 'PDB structure availability', 'enabled': True},
        'structure_quality': {'weight': 0.02, 'description': 'Structure resolution quality', 'enabled': True},
    },
    'functional_properties': {
        'functional_keywords': {'weight': 0.08, 'description': 'Functional annotation similarity', 'enabled': True},
        'organism_similarity': {'weight': 0.02, 'description': 'Organism/taxonomy similarity', 'enabled': True},
        'environment_similarity': {'weight': 0.02, 'description': 'Environment/tissue similarity', 'enabled': True},
    },
    'advanced_properties': {
        'atomic_composition': {'weight': 0.02, 'description': 'Atomic composition similarity', 'enabled': True},
    }
}

# Similarity analysis presets
SIMILARITY_PRESETS = {
    'general': {
        'name': 'General Analysis',
        'description': 'Balanced analysis across all categories',
        'weights': {
            'sequence_length': 0.10, 'sequence_identity': 0.20, 'alignment_quality': 0.10,
            'molecular_weight': 0.08, 'isoelectric_point': 0.06, 'gravy_score': 0.04,
            'extinction_coefficient': 0.02, 'full_aa_composition': 0.15,
            'hydrophobic_content': 0.05, 'charged_residues': 0.03, 'aromatic_content': 0.02,
            'polar_content': 0.02, 'structure_availability': 0.03, 'structure_quality': 0.02,
            'functional_keywords': 0.05, 'organism_similarity': 0.01, 'environment_similarity': 0.01,
            'atomic_composition': 0.01
        }
    },
    'biochemical': {
        'name': 'Biochemical Focus',
        'description': 'Emphasizes physicochemical and compositional properties',
        'weights': {
            'sequence_length': 0.05, 'sequence_identity': 0.10, 'alignment_quality': 0.05,
            'molecular_weight': 0.15, 'isoelectric_point': 0.12, 'gravy_score': 0.10,
            'extinction_coefficient': 0.08, 'full_aa_composition': 0.20,
            'hydrophobic_content': 0.03, 'charged_residues': 0.02, 'aromatic_content': 0.02,
            'polar_content': 0.02, 'structure_availability': 0.01, 'structure_quality': 0.01,
            'functional_keywords': 0.02, 'organism_similarity': 0.01, 'environment_similarity': 0.01,
            'atomic_composition': 0.01
        }
    },
    'functional': {
        'name': 'Functional Focus',
        'description': 'Emphasizes functional annotations and biological context',
        'weights': {
            'sequence_length': 0.05, 'sequence_identity': 0.10, 'alignment_quality': 0.05,
            'molecular_weight': 0.05, 'isoelectric_point': 0.03, 'gravy_score': 0.02,
            'extinction_coefficient': 0.02, 'full_aa_composition': 0.10,
            'hydrophobic_content': 0.02, 'charged_residues': 0.01, 'aromatic_content': 0.01,
            'polar_content': 0.01, 'structure_availability': 0.08, 'structure_quality': 0.05,
            'functional_keywords': 0.30, 'organism_similarity': 0.05, 'environment_similarity': 0.05,
            'atomic_composition': 0.01
        }
    }
}

def get_version_info():
    """Get formatted version information"""
    return f"{APP_NAME} v{VERSION} by {AUTHOR}"

def get_all_output_columns():
    """Get all possible output columns including optional ones"""
    all_columns = OUTPUT_COLUMNS.copy()
    all_columns.update(PDB_COLUMNS)
    all_columns.update(SIMILARITY_COLUMNS)
    return all_columns

def get_amino_acid_fields():
    """Get list of amino acid field names"""
    return list(AMINO_ACID_COLUMNS.keys())

def get_similarity_fields():
    """Get list of similarity analysis field names"""
    return list(SIMILARITY_COLUMNS.keys())

def get_required_fields():
    """Get list of required fields for basic analysis"""
    return ['organism', 'gene_name', 'function', 'sequence', 'alphafold']

def get_optional_fields():
    """Get list of optional fields"""
    all_fields = set(OUTPUT_COLUMNS.keys())
    required_fields = set(get_required_fields())
    return list(all_fields - required_fields)

def get_similarity_weights(available_dimensions):
    """Get adaptive weights for similarity analysis"""
    base_weights = SIMILARITY_CONFIG['default_weights']
    
    # Only use weights for available dimensions
    available_weights = {dim: base_weights[dim] for dim in available_dimensions if dim in base_weights}
    
    # Normalize weights to sum to 1
    total_weight = sum(available_weights.values())
    if total_weight > 0:
        return {dim: weight / total_weight for dim, weight in available_weights.items()}
    
    # Fallback: equal weights
    return {dim: 1.0 / len(available_dimensions) for dim in available_dimensions}

def categorize_similarity_zscore(z_score):
    """Categorize similarity based on Z-score"""
    thresholds = SIMILARITY_CONFIG['zscore_categories']
    
    if z_score > thresholds['highly_similar']:
        return "Highly Similar (Z>2)"
    elif z_score > thresholds['similar']:
        return "Similar (1<Z≤2)"
    elif z_score > thresholds['moderate']:
        return "Moderate (−1<Z≤1)"
    elif z_score > thresholds['distant']:
        return "Distant (−2<Z≤−1)"
    else:
        return "Unrelated (Z≤−2)"

def assess_data_quality(quality_score):
    """Assess data quality based on completeness score"""
    thresholds = SIMILARITY_CONFIG['quality_thresholds']
    
    if quality_score >= thresholds['excellent']:
        return 'Excellent'
    elif quality_score >= thresholds['good']:
        return 'Good'
    elif quality_score >= thresholds['moderate']:
        return 'Moderate'
    elif quality_score >= thresholds['poor']:
        return 'Poor'
    else:
        return 'Very Poor'

def validate_similarity_requirements(num_proteins, available_dimensions):
    """Validate if similarity analysis requirements are met"""
    config = SIMILARITY_CONFIG
    
    if num_proteins < config['min_proteins_required']:
        return False, f"Need at least {config['min_proteins_required']} proteins for similarity analysis"
    
    if len(available_dimensions) < config['min_dimensions_required']:
        return False, f"Need at least {config['min_dimensions_required']} analysis dimensions for similarity analysis"
    
    return True, "Requirements met"