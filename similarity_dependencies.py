"""
similarity_dependencies.py - Clean dependency handling for ProtMerge similarity analysis
"""

import numpy as np
import logging

logger = logging.getLogger(__name__)

# =============================================================================
# DEPENDENCY CHECKING AND IMPORTS
# =============================================================================

class DependencyManager:
    """Manages optional dependencies for similarity analysis"""
    
    def __init__(self):
        self.sklearn_available = False
        self.scipy_available = False
        self.matplotlib_available = False
        self.plotly_available = False
        
        self.cosine_similarity = None
        self.StandardScaler = None
        
        self._check_dependencies()
    
    def _check_dependencies(self):
        """Check which dependencies are available"""
        
        # Check scikit-learn
        try:
            from sklearn.metrics.pairwise import cosine_similarity as sklearn_cosine
            from sklearn.preprocessing import StandardScaler as sklearn_scaler
            self.cosine_similarity = sklearn_cosine
            self.StandardScaler = sklearn_scaler
            self.sklearn_available = True
            logger.info("scikit-learn available for similarity analysis")
        except ImportError:
            logger.warning("scikit-learn not available, using fallback methods")
            self.cosine_similarity = self._fallback_cosine_similarity
            self.StandardScaler = self._fallback_scaler
        
        # Check scipy
        try:
            import scipy
            self.scipy_available = True
            logger.info("scipy available for advanced statistics")
        except ImportError:
            logger.warning("scipy not available, some statistics limited")
        
        # Check matplotlib
        try:
            import matplotlib
            self.matplotlib_available = True
        except ImportError:
            logger.warning("matplotlib not available, some visualizations disabled")
        
        # Check plotly
        try:
            import plotly
            self.plotly_available = True
        except ImportError:
            logger.warning("plotly not available, interactive plots disabled")
    
    def _fallback_cosine_similarity(self, X, Y=None):
        """Fallback cosine similarity implementation"""
        try:
            X = np.array(X)
            if Y is None:
                Y = X
            else:
                Y = np.array(Y)
            
            # Handle single vectors
            if X.ndim == 1:
                X = X.reshape(1, -1)
            if Y.ndim == 1:
                Y = Y.reshape(1, -1)
            
            # Calculate cosine similarity manually
            X_norm = X / (np.linalg.norm(X, axis=1, keepdims=True) + 1e-8)
            Y_norm = Y / (np.linalg.norm(Y, axis=1, keepdims=True) + 1e-8)
            
            return np.dot(X_norm, Y_norm.T)
        except Exception as e:
            logger.error(f"Error in fallback cosine similarity: {e}")
            return np.array([[0.0]])
    
    def _fallback_scaler(self):
        """Fallback scaler implementation"""
        class SimpleScaler:
            def __init__(self):
                self.mean_ = None
                self.scale_ = None
            
            def fit(self, X):
                X = np.array(X)
                self.mean_ = np.mean(X, axis=0)
                self.scale_ = np.std(X, axis=0) + 1e-8
                return self
            
            def transform(self, X):
                X = np.array(X)
                if self.mean_ is None:
                    raise ValueError("Must fit scaler first")
                return (X - self.mean_) / self.scale_
            
            def fit_transform(self, X):
                return self.fit(X).transform(X)
        
        return SimpleScaler()
    
    def get_availability_report(self):
        """Get a report of available dependencies"""
        report = {
            'sklearn': self.sklearn_available,
            'scipy': self.scipy_available,
            'matplotlib': self.matplotlib_available,
            'plotly': self.plotly_available
        }
        return report
    
    def get_missing_dependencies(self):
        """Get list of missing dependencies"""
        missing = []
        if not self.sklearn_available:
            missing.append('scikit-learn')
        if not self.scipy_available:
            missing.append('scipy')
        if not self.matplotlib_available:
            missing.append('matplotlib')
        if not self.plotly_available:
            missing.append('plotly')
        return missing


# Global dependency manager instance
deps = DependencyManager()

# Export commonly used functions
cosine_similarity = deps.cosine_similarity
StandardScaler = deps.StandardScaler


# =============================================================================
# SIMILARITY CALCULATION UTILITIES
# =============================================================================

def safe_cosine_similarity(vector1, vector2):
    """Calculate cosine similarity safely with fallbacks"""
    try:
        # Ensure we have numpy arrays
        v1 = np.array(vector1, dtype=float)
        v2 = np.array(vector2, dtype=float)
        
        # Check for valid data
        if len(v1) == 0 or len(v2) == 0:
            return 0.0
        
        if len(v1) != len(v2):
            return 0.0
        
        # Remove NaN values
        valid_indices = ~(np.isnan(v1) | np.isnan(v2))
        if not np.any(valid_indices):
            return 0.0
        
        v1_clean = v1[valid_indices]
        v2_clean = v2[valid_indices]
        
        # Calculate similarity
        similarity_matrix = cosine_similarity([v1_clean], [v2_clean])
        
        # Extract scalar value
        if hasattr(similarity_matrix, 'shape'):
            if similarity_matrix.shape == (1, 1):
                return float(similarity_matrix[0, 0])
            elif similarity_matrix.shape == (1,):
                return float(similarity_matrix[0])
        
        return float(similarity_matrix)
        
    except Exception as e:
        logger.debug(f"Error in safe_cosine_similarity: {e}")
        return 0.0


def safe_ratio_similarity(value1, value2, max_diff=None):
    """Calculate ratio-based similarity safely"""
    try:
        v1 = float(value1)
        v2 = float(value2)
        
        if v1 <= 0 or v2 <= 0:
            return 0.0
        
        if max_diff is not None:
            # Normalized difference similarity
            diff = abs(v1 - v2)
            return max(0.0, 1.0 - (diff / max_diff))
        else:
            # Ratio similarity
            return min(v1, v2) / max(v1, v2)
            
    except (ValueError, TypeError, ZeroDivisionError):
        return 0.0


def safe_jaccard_similarity(set1, set2):
    """Calculate Jaccard similarity safely"""
    try:
        s1 = set(set1) if not isinstance(set1, set) else set1
        s2 = set(set2) if not isinstance(set2, set) else set2
        
        if len(s1) == 0 and len(s2) == 0:
            return 1.0
        
        intersection = len(s1.intersection(s2))
        union = len(s1.union(s2))
        
        return intersection / union if union > 0 else 0.0
        
    except Exception:
        return 0.0


def validate_protein_data(protein_data, required_fields=None):
    """Validate protein data for similarity analysis"""
    if required_fields is None:
        required_fields = ['UniProt_ID']
    
    issues = []
    
    # Check if data exists
    if protein_data is None or protein_data.empty:
        issues.append("No protein data provided")
        return False, issues
    
    # Check required fields
    for field in required_fields:
        if field not in protein_data.columns:
            issues.append(f"Missing required field: {field}")
    
    # Check minimum number of proteins
    if len(protein_data) < 3:
        issues.append(f"Need at least 3 proteins for analysis, found {len(protein_data)}")
    
    # Check for duplicate UniProt IDs
    if 'UniProt_ID' in protein_data.columns:
        duplicates = protein_data['UniProt_ID'].duplicated().sum()
        if duplicates > 0:
            issues.append(f"Found {duplicates} duplicate UniProt IDs")
    
    return len(issues) == 0, issues


def get_data_completeness(protein_data, fields_to_check):
    """Calculate data completeness for given fields"""
    completeness = {}
    
    for field in fields_to_check:
        if field not in protein_data.columns:
            completeness[field] = 0.0
            continue
        
        valid_count = 0
        total_count = len(protein_data)
        
        for value in protein_data[field]:
            if is_valid_value(value):
                valid_count += 1
        
        completeness[field] = valid_count / total_count if total_count > 0 else 0.0
    
    return completeness


def is_valid_value(value):
    """Check if a value is valid (not missing or placeholder)"""
    if value is None or np.isnan(value) if isinstance(value, (int, float)) else False:
        return False
    
    value_str = str(value).strip().upper()
    invalid_values = ['', 'NO VALUE FOUND', 'NAN', 'NONE', 'N/A', 'UNKNOWN', 'NULL']
    
    return value_str not in invalid_values


# =============================================================================
# SIMILARITY PRESETS WITH FALLBACKS
# =============================================================================

class RobustSimilarityPresets:
    """Similarity presets that adapt to available data and dependencies"""
    
    @staticmethod
    def get_available_categories(protein_data):
        """Get categories that have sufficient data"""
        if protein_data is None or protein_data.empty:
            return {}
        
        available = {}
        
        # Sequence properties
        if 'sequence' in protein_data.columns:
            seq_completeness = get_data_completeness(protein_data, ['sequence'])['sequence']
            if seq_completeness > 0.5:  # At least 50% have sequences
                available['sequence_length'] = 0.15
                available['sequence_identity'] = 0.10
        
        # Physicochemical properties
        physico_fields = ['mw', 'pi', 'gravy', 'ext']
        physico_completeness = get_data_completeness(protein_data, physico_fields)
        
        if physico_completeness.get('mw', 0) > 0.3:
            available['molecular_weight'] = 0.20
        if physico_completeness.get('pi', 0) > 0.3:
            available['isoelectric_point'] = 0.15
        if physico_completeness.get('gravy', 0) > 0.3:
            available['gravy_score'] = 0.10
        if physico_completeness.get('ext', 0) > 0.3:
            available['extinction_coefficient'] = 0.05
        
        # Functional properties
        if 'keywords' in protein_data.columns:
            kw_completeness = get_data_completeness(protein_data, ['keywords'])['keywords']
            if kw_completeness > 0.3:
                available['functional_keywords'] = 0.15
        
        if 'organism' in protein_data.columns:
            org_completeness = get_data_completeness(protein_data, ['organism'])['organism']
            if org_completeness > 0.5:
                available['organism_similarity'] = 0.10
        
        return available
    
    @staticmethod
    def get_adaptive_weights(protein_data):
        """Get weights adapted to available data"""
        available = RobustSimilarityPresets.get_available_categories(protein_data)
        
        if not available:
            # Fallback to minimal weights
            return {'sequence_length': 1.0}
        
        # Normalize weights to sum to 1
        total_weight = sum(available.values())
        if total_weight > 0:
            return {k: v / total_weight for k, v in available.items()}
        
        return available


# =============================================================================
# EXPORT FUNCTIONS
# =============================================================================

__all__ = [
    'deps',
    'cosine_similarity', 
    'StandardScaler',
    'safe_cosine_similarity',
    'safe_ratio_similarity', 
    'safe_jaccard_similarity',
    'validate_protein_data',
    'get_data_completeness',
    'is_valid_value',
    'RobustSimilarityPresets'
]