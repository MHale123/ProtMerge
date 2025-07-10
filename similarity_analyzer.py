"""
similarity_analyzer.py - Complete Rewrite for ProtMerge v1.2.0
"""

import pandas as pd
import numpy as np  # FIXED: Added missing numpy import - this was causing crashes
import re
import logging
import math
from typing import Dict, List, Tuple, Optional, Any


class SimilarityAnalyzer:
    """
    Standalone similarity analyzer that works without external dependencies.
    Provides robust protein similarity analysis using built-in methods.
    """
    
    def __init__(self):
        self.name = "Similarity"
        self.logger = logging.getLogger(f"{__name__}.{self.name}")
        
        # Core data storage
        self.protein_data = None
        self.precomputed_scores = {}
        self.data_quality_scores = {}
        
        # Available similarity functions
        self.similarity_functions = {
            'sequence_length': self._calc_sequence_length_similarity,
            'molecular_weight': self._calc_molecular_weight_similarity,
            'isoelectric_point': self._calc_isoelectric_point_similarity,
            'gravy_score': self._calc_gravy_similarity,
            'sequence_identity': self._calc_sequence_identity_similarity,
            'functional_keywords': self._calc_functional_keywords_similarity,
            'organism_similarity': self._calc_organism_similarity,
            'extinction_coefficient': self._calc_extinction_coefficient_similarity,
            'amino_acid_composition': self._calc_amino_acid_similarity,
        }
        
        self.logger.info("SimilarityAnalyzer initialized successfully")
    
    def analyze(self, results, options, progress_callback=None):
        """
        Pre-compute all similarity scores for protein pairs.
        
        Args:
            results: DataFrame with protein data
            options: Analysis options (not used but kept for compatibility)
            progress_callback: Function to report progress
        """
        self.logger.info("Starting similarity analysis pre-computation")
        
        # Validate input
        if results is None or results.empty:
            raise ValueError("No protein data provided")
        
        if len(results) < 2:
            raise ValueError(f"Need at least 2 proteins for similarity analysis, got {len(results)}")
        
        self.protein_data = results.copy()
        protein_ids = results['UniProt_ID'].tolist()
        
        self.logger.info(f"Analyzing {len(protein_ids)} proteins")
        
        # Calculate data quality scores
        self._calculate_data_quality_scores()
        
        # Pre-compute pairwise similarities
        total_pairs = len(protein_ids) * (len(protein_ids) - 1) // 2
        computed_pairs = 0
        
        self.logger.info(f"Computing {total_pairs} protein pair similarities")
        
        for i, protein1 in enumerate(protein_ids):
            for j, protein2 in enumerate(protein_ids):
                if i >= j:  # Only compute upper triangle
                    continue
                
                pair_key = (protein1, protein2)
                
                try:
                    scores = self._compute_pairwise_similarities(protein1, protein2)
                    self.precomputed_scores[pair_key] = scores
                except Exception as e:
                    self.logger.warning(f"Error computing similarity for {protein1}-{protein2}: {e}")
                    # Store zeros for failed computations
                    self.precomputed_scores[pair_key] = {func: 0.0 for func in self.similarity_functions.keys()}
                
                computed_pairs += 1
                
                # Progress reporting
                if progress_callback and computed_pairs % max(1, total_pairs // 20) == 0:
                    progress = (computed_pairs / total_pairs) * 90  # Reserve 10% for final processing
                    progress_callback(progress, f"Computing similarities ({computed_pairs}/{total_pairs})")
        
        if progress_callback:
            progress_callback(100, "Similarity pre-computation complete")
        
        self.logger.info(f"Successfully computed {computed_pairs} protein pair similarities")
    
    def calculate_similarity_matrix(self, central_protein_id: str, weights: Dict[str, float]) -> pd.DataFrame:
        """
        Calculate similarity matrix relative to a central protein.
        
        Args:
            central_protein_id: ID of protein to compare all others against
            weights: Dictionary of category weights
        
        Returns:
            DataFrame with similarity results sorted by overall similarity
        """
        if self.protein_data is None:
            raise ValueError("Must run analyze() first")
        
        protein_ids = self.protein_data['UniProt_ID'].tolist()
        
        if central_protein_id not in protein_ids:
            raise ValueError(f"Central protein {central_protein_id} not found in dataset")
        
        self.logger.info(f"Calculating similarity matrix for central protein: {central_protein_id}")
        
        # Normalize weights
        valid_weights = {k: float(v) for k, v in weights.items() if isinstance(v, (int, float)) and v > 0}
        if not valid_weights:
            # Default equal weights for available functions
            valid_weights = {func: 1.0 for func in self.similarity_functions.keys()}
        
        total_weight = sum(valid_weights.values())
        if total_weight > 0:
            valid_weights = {k: v/total_weight for k, v in valid_weights.items()}
        
        # Calculate similarities
        similarities = []
        
        for protein_id in protein_ids:
            if protein_id == central_protein_id:
                continue
            
            try:
                # Get precomputed scores
                pair_key = (min(central_protein_id, protein_id), max(central_protein_id, protein_id))
                
                if pair_key not in self.precomputed_scores:
                    self.logger.warning(f"No precomputed scores for {pair_key}")
                    continue
                
                scores = self.precomputed_scores[pair_key]
                
                # Calculate weighted overall similarity
                overall_similarity = self._calculate_weighted_similarity(scores, valid_weights)
                
                # Create result record
                result = {
                    'protein_id': protein_id,
                    'overall_similarity': overall_similarity,
                    'data_quality': self.data_quality_scores.get(protein_id, 0.0)
                }
                
                # Add individual category scores
                for category, score in scores.items():
                    result[category] = score
                
                similarities.append(result)
                
            except Exception as e:
                self.logger.warning(f"Error processing {protein_id}: {e}")
                continue
        
        # Convert to DataFrame and sort
        if similarities:
            df = pd.DataFrame(similarities)
            df = df.sort_values('overall_similarity', ascending=False)
            self.logger.info(f"Generated similarity matrix with {len(df)} proteins")
            return df
        else:
            self.logger.warning("No similarity results generated")
            return pd.DataFrame()
    
    def _compute_pairwise_similarities(self, protein1: str, protein2: str) -> Dict[str, float]:
        """Compute all similarity scores between two proteins."""
        scores = {}
        
        for category, func in self.similarity_functions.items():
            try:
                score = func(protein1, protein2)
                # Ensure valid score
                if score is None or (isinstance(score, float) and math.isnan(score)):
                    scores[category] = 0.0
                else:
                    scores[category] = max(0.0, min(1.0, float(score)))
            except Exception as e:
                self.logger.debug(f"Error in {category} for {protein1}-{protein2}: {e}")
                scores[category] = 0.0
        
        return scores
    
    def _calculate_weighted_similarity(self, scores: Dict[str, float], weights: Dict[str, float]) -> float:
        """Calculate weighted average similarity."""
        total_score = 0.0
        total_weight = 0.0
        
        for category, score in scores.items():
            weight = weights.get(category, 0.0)
            if weight > 0 and isinstance(score, (int, float)):
                total_score += score * weight
                total_weight += weight
        
        return total_score / total_weight if total_weight > 0 else 0.0
    
    def _get_protein_data(self, protein_id: str) -> pd.Series:
        """Get data for a specific protein."""
        mask = self.protein_data['UniProt_ID'] == protein_id
        matches = self.protein_data[mask]
        
        if matches.empty:
            raise ValueError(f"Protein {protein_id} not found")
        
        return matches.iloc[0]
    
    def _is_valid_value(self, value) -> bool:
        """Check if a value is valid (not missing or placeholder)."""
        if value is None or pd.isna(value):
            return False
        
        value_str = str(value).strip().upper()
        invalid_values = {'', 'NO VALUE FOUND', 'NAN', 'NONE', 'N/A', 'UNKNOWN', 'NULL'}
        
        return value_str not in invalid_values
    
    def _calculate_data_quality_scores(self):
        """Calculate data completeness score for each protein."""
        self.data_quality_scores = {}
        
        # Fields to check for quality assessment
        quality_fields = ['sequence', 'mw', 'pi', 'gravy', 'ext', 'function', 'keywords', 'organism']
        
        for _, row in self.protein_data.iterrows():
            protein_id = row['UniProt_ID']
            available_count = 0
            
            for field in quality_fields:
                if field in row and self._is_valid_value(row[field]):
                    available_count += 1
            
            quality_score = available_count / len(quality_fields)
            self.data_quality_scores[protein_id] = quality_score
    
    # =============================================================================
    # SIMILARITY CALCULATION FUNCTIONS
    # =============================================================================
    
    def _calc_sequence_length_similarity(self, p1: str, p2: str) -> float:
        """Compare sequence lengths using ratio method."""
        try:
            data1 = self._get_protein_data(p1)
            data2 = self._get_protein_data(p2)
            
            seq1 = data1.get('sequence', '')
            seq2 = data2.get('sequence', '')
            
            if not self._is_valid_value(seq1) or not self._is_valid_value(seq2):
                return 0.0
            
            len1, len2 = len(str(seq1)), len(str(seq2))
            if len1 == 0 or len2 == 0:
                return 0.0
            
            return min(len1, len2) / max(len1, len2)
        except Exception:
            return 0.0
    
    def _calc_molecular_weight_similarity(self, p1: str, p2: str) -> float:
        """Compare molecular weights using ratio method."""
        try:
            data1 = self._get_protein_data(p1)
            data2 = self._get_protein_data(p2)
            
            mw1 = data1.get('mw', None)
            mw2 = data2.get('mw', None)
            
            if not self._is_valid_value(mw1) or not self._is_valid_value(mw2):
                return 0.0
            
            mw1_val, mw2_val = float(mw1), float(mw2)
            if mw1_val <= 0 or mw2_val <= 0:
                return 0.0
            
            return min(mw1_val, mw2_val) / max(mw1_val, mw2_val)
        except (ValueError, TypeError):
            return 0.0
    
    def _calc_isoelectric_point_similarity(self, p1: str, p2: str) -> float:
        """Compare isoelectric points normalized over pH scale."""
        try:
            data1 = self._get_protein_data(p1)
            data2 = self._get_protein_data(p2)
            
            pi1 = data1.get('pi', None)
            pi2 = data2.get('pi', None)
            
            if not self._is_valid_value(pi1) or not self._is_valid_value(pi2):
                return 0.0
            
            pi1_val, pi2_val = float(pi1), float(pi2)
            diff = abs(pi1_val - pi2_val)
            
            # Normalize over pH scale (0-14)
            return max(0.0, 1.0 - (diff / 14.0))
        except (ValueError, TypeError):
            return 0.0
    
    def _calc_gravy_similarity(self, p1: str, p2: str) -> float:
        """Compare GRAVY scores normalized over typical range."""
        try:
            data1 = self._get_protein_data(p1)
            data2 = self._get_protein_data(p2)
            
            gravy1 = data1.get('gravy', None)
            gravy2 = data2.get('gravy', None)
            
            if not self._is_valid_value(gravy1) or not self._is_valid_value(gravy2):
                return 0.0
            
            gravy1_val, gravy2_val = float(gravy1), float(gravy2)
            diff = abs(gravy1_val - gravy2_val)
            
            # GRAVY typically ranges from -2 to +2
            return max(0.0, 1.0 - (diff / 4.0))
        except (ValueError, TypeError):
            return 0.0
    
    def _calc_sequence_identity_similarity(self, p1: str, p2: str) -> float:
        """Compare sequence identity from BLAST data."""
        try:
            data1 = self._get_protein_data(p1)
            data2 = self._get_protein_data(p2)
            
            identity1 = data1.get('identity', None)
            identity2 = data2.get('identity', None)
            
            if not self._is_valid_value(identity1) or not self._is_valid_value(identity2):
                return 0.0
            
            id1_val, id2_val = float(identity1), float(identity2)
            diff = abs(id1_val - id2_val)
            
            # Normalize over 0-100% range
            return max(0.0, 1.0 - (diff / 100.0))
        except (ValueError, TypeError):
            return 0.0
    
    def _calc_functional_keywords_similarity(self, p1: str, p2: str) -> float:
        """Compare functional keywords using Jaccard similarity."""
        try:
            data1 = self._get_protein_data(p1)
            data2 = self._get_protein_data(p2)
            
            kw1 = data1.get('keywords', '')
            kw2 = data2.get('keywords', '')
            
            if not self._is_valid_value(kw1) or not self._is_valid_value(kw2):
                return 0.0
            
            # Parse keywords
            set1 = set(kw.strip().lower() for kw in str(kw1).split(';') if kw.strip())
            set2 = set(kw.strip().lower() for kw in str(kw2).split(';') if kw.strip())
            
            if not set1 or not set2:
                return 0.0
            
            # Jaccard similarity
            intersection = len(set1.intersection(set2))
            union = len(set1.union(set2))
            
            return intersection / union if union > 0 else 0.0
        except Exception:
            return 0.0
    
    def _calc_organism_similarity(self, p1: str, p2: str) -> float:
        """Compare organism similarity."""
        try:
            data1 = self._get_protein_data(p1)
            data2 = self._get_protein_data(p2)
            
            org1 = data1.get('organism', '')
            org2 = data2.get('organism', '')
            
            if not self._is_valid_value(org1) or not self._is_valid_value(org2):
                return 0.0
            
            org1_str = str(org1).lower().strip()
            org2_str = str(org2).lower().strip()
            
            # Exact match
            if org1_str == org2_str:
                return 1.0
            
            # Genus match (first word)
            try:
                genus1 = org1_str.split()[0]
                genus2 = org2_str.split()[0]
                return 0.5 if genus1 == genus2 else 0.0
            except IndexError:
                return 0.0
        except Exception:
            return 0.0
    
    def _calc_extinction_coefficient_similarity(self, p1: str, p2: str) -> float:
        """Compare extinction coefficients using ratio method."""
        try:
            data1 = self._get_protein_data(p1)
            data2 = self._get_protein_data(p2)
            
            ext1 = data1.get('ext', None)
            ext2 = data2.get('ext', None)
            
            if not self._is_valid_value(ext1) or not self._is_valid_value(ext2):
                return 0.0
            
            ext1_val, ext2_val = float(ext1), float(ext2)
            if ext1_val <= 0 or ext2_val <= 0:
                return 0.0
            
            return min(ext1_val, ext2_val) / max(ext1_val, ext2_val)
        except (ValueError, TypeError):
            return 0.0
    
    def _calc_amino_acid_similarity(self, p1: str, p2: str) -> float:
        """Compare amino acid composition using simple correlation."""
        try:
            comp1 = self._get_amino_acid_vector(p1)
            comp2 = self._get_amino_acid_vector(p2)
            
            if comp1 is None or comp2 is None:
                return 0.0
            
            # Simple correlation-based similarity
            return self._calculate_vector_similarity(comp1, comp2)
        except Exception:
            return 0.0
    
    def _get_amino_acid_vector(self, protein_id: str) -> Optional[List[float]]:
        """Get amino acid composition as percentage vector."""
        try:
            data = self._get_protein_data(protein_id)
            
            aa_percentages = []
            aa_keys = ['ala', 'arg', 'asn', 'asp', 'cys', 'gln', 'glu', 'gly', 
                      'his', 'ile', 'leu', 'lys', 'met', 'phe', 'pro', 'ser', 
                      'thr', 'trp', 'tyr', 'val']
            
            for aa_key in aa_keys:
                aa_value = data.get(aa_key, '0_0.0%')
                
                if self._is_valid_value(aa_value):
                    try:
                        # Extract percentage from format "20_12.3%"
                        value_str = str(aa_value)
                        if '_' in value_str and '%' in value_str:
                            percentage = float(value_str.split('_')[1].rstrip('%'))
                            aa_percentages.append(percentage)
                        else:
                            aa_percentages.append(0.0)
                    except (IndexError, ValueError):
                        aa_percentages.append(0.0)
                else:
                    aa_percentages.append(0.0)
            
            # Return only if we have some data
            return aa_percentages if any(p > 0 for p in aa_percentages) else None
        except Exception:
            return None
    
    def _calculate_vector_similarity(self, vec1: List[float], vec2: List[float]) -> float:
        """Calculate similarity between two vectors using cosine similarity."""
        try:
            # Convert to numpy arrays
            v1 = np.array(vec1, dtype=float)
            v2 = np.array(vec2, dtype=float)
            
            # Check for valid data
            if len(v1) != len(v2) or len(v1) == 0:
                return 0.0
            
            # Calculate norms
            norm1 = np.linalg.norm(v1)
            norm2 = np.linalg.norm(v2)
            
            if norm1 == 0 or norm2 == 0:
                return 0.0
            
            # Cosine similarity
            similarity = np.dot(v1, v2) / (norm1 * norm2)
            return max(0.0, float(similarity))
        except Exception:
            return 0.0
    
    def get_available_categories(self) -> Dict[str, str]:
        """Get available similarity categories with descriptions."""
        return {
            'sequence_length': 'Sequence length similarity',
            'molecular_weight': 'Molecular weight similarity',
            'isoelectric_point': 'Isoelectric point similarity',
            'gravy_score': 'Hydrophobicity (GRAVY) similarity',
            'sequence_identity': 'BLAST sequence identity similarity',
            'functional_keywords': 'Functional annotation similarity',
            'organism_similarity': 'Organism/species similarity',
            'extinction_coefficient': 'Extinction coefficient similarity',
            'amino_acid_composition': 'Amino acid composition similarity'
        }


# =============================================================================
# PRESET CONFIGURATIONS
# =============================================================================

class SimilarityPresets:
    """Predefined similarity analysis configurations."""
    
    @staticmethod
    def get_basic_preset():
        """Basic preset focusing on core properties."""
        return {
            'sequence_length': 0.25,
            'molecular_weight': 0.25,
            'isoelectric_point': 0.25,
            'gravy_score': 0.25
        }
    
    @staticmethod
    def get_sequence_preset():
        """Sequence-focused preset."""
        return {
            'sequence_length': 0.30,
            'sequence_identity': 0.30,
            'amino_acid_composition': 0.25,
            'molecular_weight': 0.15
        }
    
    @staticmethod
    def get_biochemical_preset():
        """Biochemical properties preset."""
        return {
            'molecular_weight': 0.30,
            'isoelectric_point': 0.25,
            'gravy_score': 0.20,
            'extinction_coefficient': 0.15,
            'sequence_length': 0.10
        }
    
    @staticmethod
    def get_functional_preset():
        """Functional annotation preset."""
        return {
            'functional_keywords': 0.40,
            'organism_similarity': 0.25,
            'sequence_identity': 0.20,
            'molecular_weight': 0.15
        }
    
    @staticmethod
    def adapt_weights_to_data(protein_data):
        """Adapt weights based on data availability."""
        if protein_data is None or protein_data.empty:
            return SimilarityPresets.get_basic_preset()
        
        # Check data completeness
        weights = {}
        weight_sum = 0.0
        
        # Core properties (always try to include)
        if 'sequence' in protein_data.columns:
            completeness = sum(1 for v in protein_data['sequence'] if pd.notna(v) and str(v) != 'NO VALUE FOUND') / len(protein_data)
            if completeness > 0.3:
                weights['sequence_length'] = 0.2
                weight_sum += 0.2
        
        if 'mw' in protein_data.columns:
            completeness = sum(1 for v in protein_data['mw'] if pd.notna(v) and str(v) != 'NO VALUE FOUND') / len(protein_data)
            if completeness > 0.3:
                weights['molecular_weight'] = 0.25
                weight_sum += 0.25
        
        if 'pi' in protein_data.columns:
            completeness = sum(1 for v in protein_data['pi'] if pd.notna(v) and str(v) != 'NO VALUE FOUND') / len(protein_data)
            if completeness > 0.3:
                weights['isoelectric_point'] = 0.2
                weight_sum += 0.2
        
        if 'gravy' in protein_data.columns:
            completeness = sum(1 for v in protein_data['gravy'] if pd.notna(v) and str(v) != 'NO VALUE FOUND') / len(protein_data)
            if completeness > 0.3:
                weights['gravy_score'] = 0.15
                weight_sum += 0.15
        
        # Optional properties
        if 'keywords' in protein_data.columns:
            completeness = sum(1 for v in protein_data['keywords'] if pd.notna(v) and str(v) != 'NO VALUE FOUND') / len(protein_data)
            if completeness > 0.3:
                weights['functional_keywords'] = 0.1
                weight_sum += 0.1
        
        if 'organism' in protein_data.columns:
            completeness = sum(1 for v in protein_data['organism'] if pd.notna(v) and str(v) != 'NO VALUE FOUND') / len(protein_data)
            if completeness > 0.5:
                weights['organism_similarity'] = 0.1
                weight_sum += 0.1
        
        # Normalize weights
        if weight_sum > 0:
            weights = {k: v/weight_sum for k, v in weights.items()}
            return weights
        else:
            # Fallback to basic preset
            return SimilarityPresets.get_basic_preset()


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def run_similarity_analysis(results_df, central_protein_id, weights=None, progress_callback=None):
    """
    Convenience function to run complete similarity analysis.
    
    Args:
        results_df: DataFrame with protein data
        central_protein_id: ID of central protein
        weights: Similarity weights (optional)
        progress_callback: Progress callback function
    
    Returns:
        DataFrame with similarity results
    """
    try:
        # Initialize analyzer
        analyzer = SimilarityAnalyzer()
        
        # Use adaptive weights if none provided
        if weights is None:
            weights = SimilarityPresets.adapt_weights_to_data(results_df)
        
        # Run analysis
        analyzer.analyze(results_df, {}, progress_callback)
        
        # Calculate similarity matrix
        similarity_results = analyzer.calculate_similarity_matrix(central_protein_id, weights)
        
        return similarity_results
        
    except Exception as e:
        logging.error(f"Similarity analysis failed: {e}")
        raise


def get_similarity_summary(results_df):
    """Get summary of available data for similarity analysis."""
    if results_df is None or results_df.empty:
        return "No data available"
    
    total_proteins = len(results_df)
    
    # Check key fields
    fields_to_check = ['sequence', 'mw', 'pi', 'gravy', 'keywords', 'organism']
    availability = {}
    
    for field in fields_to_check:
        if field in results_df.columns:
            valid_count = sum(1 for v in results_df[field] 
                            if pd.notna(v) and str(v) not in ['', 'NO VALUE FOUND', 'nan'])
            availability[field] = f"{valid_count}/{total_proteins} ({valid_count/total_proteins*100:.1f}%)"
        else:
            availability[field] = "Not available"
    
    summary = f"Protein Data Summary ({total_proteins} proteins):\n"
    for field, status in availability.items():
        summary += f"  {field}: {status}\n"
    
    return summary