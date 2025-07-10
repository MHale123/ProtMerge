"""
Protein analysis modules for ProtMerge v1.2.0
"""

import requests
import time
import re
import xml.etree.ElementTree as ET
import logging
import pandas as pd
from config import *

class AnalyzerManager:
    """Manages all protein analyzers"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.analyzers = {
            'uniprot': UniProtAnalyzer(),
            'protparam': ProtParamAnalyzer(),
            'blast': BLASTAnalyzer(),
            'pdb': PDBAnalyzer()
            # Note: similarity analyzer is loaded on-demand
        }
        self.logger.info("AnalyzerManager initialized")
    
    def run_all_analyses(self, data, options, progress_callback=None):
        """Run all selected analyses"""
        results = data['results']
        
        if options.get('uniprot', True):
            if progress_callback:
                progress_callback(10, "Fetching UniProt data", "Retrieving protein information")
            self.analyzers['uniprot'].analyze(results, options, progress_callback)
        
        if options.get('protparam', False):
            if progress_callback:
                progress_callback(50, "Running ProtParam analysis", "Calculating molecular properties")
            self.analyzers['protparam'].analyze(results, options, progress_callback)
        
        if options.get('blast', False):
            if progress_callback:
                progress_callback(75, "Running BLAST analysis", "Searching for similar proteins")
            self.analyzers['blast'].analyze(results, options, progress_callback)
        
        if options.get('pdb_search', False):
            if progress_callback:
                progress_callback(85, "Searching PDB structures", "Finding structural information")
            self.analyzers['pdb'].analyze(results, options, progress_callback)
        
        return results
    
    def run_similarity_analysis(self, data, central_protein_id, custom_weights, progress_callback=None):
        """Run similarity analysis with custom parameters"""
        try:
            if progress_callback:
                progress_callback(0, "Initializing similarity analysis")
            
            # Import similarity analyzer on-demand to avoid circular imports
            try:
                from similarity_analyzer import SimilarityAnalyzer
            except ImportError as e:
                self.logger.error(f"Failed to import similarity analyzer: {e}")
                raise ImportError("Similarity analysis not available. Please ensure similarity_analyzer.py is present.")
            
            # Initialize similarity analyzer if not already done
            if 'similarity' not in self.analyzers:
                self.analyzers['similarity'] = SimilarityAnalyzer()
            
            # First run pre-computation if not already done
            if not hasattr(self.analyzers['similarity'], 'precomputed_scores') or \
               not self.analyzers['similarity'].precomputed_scores:
                if progress_callback:
                    progress_callback(10, "Pre-computing similarity scores")
                self.analyzers['similarity'].analyze(data['results'], {}, progress_callback)
            
            if progress_callback:
                progress_callback(80, "Calculating similarity matrix")
            
            # Calculate similarity matrix
            similarity_results = self.analyzers['similarity'].calculate_similarity_matrix(
                central_protein_id, custom_weights
            )
            
            if progress_callback:
                progress_callback(100, "Similarity analysis complete")
            
            return similarity_results
            
        except Exception as e:
            self.logger.error(f"Similarity analysis failed: {e}")
            raise
        
class BaseAnalyzer:
    """Base class for all analyzers"""
    
    def __init__(self, name):
        self.name = name
        self.logger = logging.getLogger(f"{__name__}.{name}")
    
    def make_request(self, url, method='GET', **kwargs):
        """Make HTTP request with rate limiting"""
        try:
            if method.upper() == 'GET':
                response = requests.get(url, timeout=30, **kwargs)
            elif method.upper() == 'POST':
                response = requests.post(url, timeout=30, **kwargs)
            else:
                raise ValueError(f"Unsupported HTTP method: {method}")
            
            time.sleep(RATE_LIMITS.get(self.name.lower(), 0.1))
            return response
            
        except Exception as e:
            self.logger.error(f"{self.name} request failed: {e}")
            raise
    
    def should_update(self, results, idx, field, safe_mode=True):
        """Check if field should be updated"""
        if not safe_mode:
            return True
        current_value = results.at[idx, field]
        return (pd.isna(current_value) or 
                str(current_value).strip() == '' or 
                str(current_value) == "NO VALUE FOUND")
    
    def set_no_value(self, results, idx, fields, safe_mode=True):
        """Set multiple fields to NO VALUE FOUND"""
        for field in fields:
            if self.should_update(results, idx, field, safe_mode):
                results.at[idx, field] = "NO VALUE FOUND"

class UniProtAnalyzer(BaseAnalyzer):
    """UniProt database analyzer"""
    
    def __init__(self):
        super().__init__("UniProt")
    
    def analyze(self, results, options, progress_callback=None):
        """Run UniProt analysis"""
        self.logger.info("Starting UniProt analysis")
        
        safe_mode = options.get('safe_mode', True)
        to_process = self._get_processing_list(results, safe_mode)
        
        if not to_process:
            self.logger.info("All UniProt data complete")
            return
        
        self.logger.info(f"Processing {len(to_process)} entries")
        
        for i, (idx, uniprot_id) in enumerate(to_process):
            if progress_callback:
                progress = 10 + (40 * (i + 1) / len(to_process))
                progress_callback(progress, f"UniProt ({i+1}/{len(to_process)})", f"Processing {uniprot_id}")
            
            try:
                url = f"{UNIPROT_BASE_URL}/{uniprot_id}.json"
                response = self.make_request(url)
                
                if response.status_code == 200:
                    self._process_data(results, idx, response.json(), safe_mode)
                else:
                    self._set_no_value(results, idx, safe_mode)
                    
            except Exception as e:
                self.logger.error(f"Error processing {uniprot_id}: {e}")
                self._set_no_value(results, idx, safe_mode)
    
    def _get_processing_list(self, results, safe_mode):
        """Get entries that need processing"""
        to_process = []
        fields = ['organism', 'gene_name', 'function', 'sequence', 'environment', 'keywords', 'structure', 'alphafold']
        
        for idx, row in results.iterrows():
            if any(self.should_update(results, idx, field, safe_mode) for field in fields):
                to_process.append((idx, row['UniProt_ID']))
        
        return to_process
    
    def _process_data(self, results, idx, data, safe_mode):
        """Process UniProt JSON data"""
        if self.should_update(results, idx, 'organism', safe_mode):
            organism = data.get("organism", {}).get("scientificName", "NO VALUE FOUND")
            results.at[idx, 'organism'] = organism
        
        if self.should_update(results, idx, 'gene_name', safe_mode):
            gene_name = "NO VALUE FOUND"
            genes = data.get("genes", [])
            if genes and "geneName" in genes[0]:
                gene_name = genes[0]["geneName"].get("value", "NO VALUE FOUND")
            results.at[idx, 'gene_name'] = gene_name
        
        if self.should_update(results, idx, 'function', safe_mode):
            function = "NO VALUE FOUND"
            for comment in data.get("comments", []):
                if comment.get("commentType") == "FUNCTION":
                    texts = comment.get("texts", [])
                    if texts:
                        function = texts[0].get("value", "NO VALUE FOUND")
                        break
            results.at[idx, 'function'] = function
        
        if self.should_update(results, idx, 'sequence', safe_mode):
            sequence = data.get("sequence", {}).get("value", "NO VALUE FOUND")
            results.at[idx, 'sequence'] = sequence
        
        if self.should_update(results, idx, 'environment', safe_mode):
            environment = self._extract_environment(data)
            results.at[idx, 'environment'] = environment
        
        if self.should_update(results, idx, 'keywords', safe_mode):
            keywords = []
            for kw in data.get("keywords", []):
                keyword_name = kw.get("name", "")
                if keyword_name:
                    keywords.append(keyword_name)
            results.at[idx, 'keywords'] = "; ".join(keywords) if keywords else "NO VALUE FOUND"
        
        if self.should_update(results, idx, 'structure', safe_mode):
            features = []
            for feature in data.get("features", []):
                ftype = feature.get("type", "")
                if ftype in ['Domain', 'Region', 'Transmembrane', 'Signal', 'Topological domain']:
                    features.append(ftype)
            unique_features = list(set(features))
            results.at[idx, 'structure'] = "; ".join(unique_features) if unique_features else "NO VALUE FOUND"
        
        if self.should_update(results, idx, 'alphafold', safe_mode):
            uniprot_id = results.at[idx, 'UniProt_ID']
            results.at[idx, 'alphafold'] = f"https://alphafold.ebi.ac.uk/entry/{uniprot_id}"
    
    def _extract_environment(self, data):
        """Extract body location/environment"""
        locations = []
        text_sources = []
        
        for comment in data.get("comments", []):
            if comment.get("commentType") in ["SUBCELLULAR LOCATION", "FUNCTION"]:
                texts = comment.get("texts", [])
                if texts:
                    text_sources.append(texts[0].get("value", "").lower())
        
        organism = data.get("organism", {}).get("scientificName", "").lower()
        text_sources.append(organism)
        
        for kw in data.get("keywords", []):
            text_sources.append(kw.get("name", "").lower())
        
        all_text = " ".join(text_sources)
        for body_part, keywords in BODY_LOCATIONS.items():
            if any(keyword in all_text for keyword in keywords):
                locations.append(body_part)
        
        for bacteria in GUT_BACTERIA:
            if bacteria in organism:
                locations.append('gut')
        
        unique_locations = list(set(locations))
        if unique_locations:
            return "; ".join(sorted(unique_locations))
        
        organism_name = data.get("organism", {}).get("scientificName", "")
        return f"Found in: {organism_name}" if organism_name else "NO VALUE FOUND"
    
    def _set_no_value(self, results, idx, safe_mode):
        """Set all UniProt fields to NO VALUE FOUND"""
        fields = ['organism', 'gene_name', 'function', 'sequence', 'environment', 'keywords', 'structure']
        self.set_no_value(results, idx, fields, safe_mode)
        
        if self.should_update(results, idx, 'alphafold', safe_mode):
            uniprot_id = results.at[idx, 'UniProt_ID']
            results.at[idx, 'alphafold'] = f"https://alphafold.ebi.ac.uk/entry/{uniprot_id}"

class ProtParamAnalyzer(BaseAnalyzer):
    """ProtParam analyzer for protein properties"""
    
    def __init__(self):
        super().__init__("ProtParam")
    
    def analyze(self, results, options, progress_callback=None):
        """Run ProtParam analysis"""
        self.logger.info("Starting ProtParam analysis")
        
        safe_mode = options.get('safe_mode', True)
        to_process = self._get_processing_list(results, options, safe_mode)
        
        if not to_process:
            self.logger.info("All ProtParam data complete")
            return
        
        self.logger.info(f"Processing {len(to_process)} sequences")
        
        for i, (idx, sequence) in enumerate(to_process):
            if progress_callback:
                progress = 50 + (25 * (i + 1) / len(to_process))
                progress_callback(progress, f"ProtParam ({i+1}/{len(to_process)})", f"Analyzing sequence {i+1}")
            
            clean_seq = re.sub(r'[^ACDEFGHIKLMNPQRSTVWY]', '', sequence.upper())
            
            if len(clean_seq) < 20:
                self._set_no_value(results, idx, options, safe_mode)
                continue
            
            try:
                response = self._submit_protparam(clean_seq)
                
                if response and response.status_code == 200:
                    self._parse_response(results, idx, response.text, options, safe_mode)
                else:
                    self._set_no_value(results, idx, options, safe_mode)
                    
            except Exception as e:
                self.logger.error(f"ProtParam error: {e}")
                self._set_no_value(results, idx, options, safe_mode)
    
    def _get_processing_list(self, results, options, safe_mode):
        """Get sequences that need ProtParam processing"""
        to_process = []
        protparam_fields = ['mw', 'pi', 'gravy', 'ext']
        if options.get('amino_acid', False):
            protparam_fields.extend(AMINO_ACID_COLUMNS.keys())
        
        for idx, row in results.iterrows():
            sequence = row.get('sequence', '')
            if sequence and sequence != "NO VALUE FOUND" and len(sequence) >= 20:
                if any(self.should_update(results, idx, field, safe_mode) for field in protparam_fields):
                    to_process.append((idx, sequence))
            else:
                self._set_no_value(results, idx, options, safe_mode)
        
        return to_process
    
    def _submit_protparam(self, sequence):
        """Submit sequence to ProtParam web service"""
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        
        data = {'sequence': sequence}
        response = self.make_request(PROTPARAM_URL, method='POST', data=data, headers=headers)
        time.sleep(1.5)
        return response
    
    def _parse_response(self, results, idx, html, options, safe_mode):
        """Parse ProtParam HTML response"""
        for param_key in ['mw', 'pi', 'gravy', 'ext']:
            if self.should_update(results, idx, param_key, safe_mode):
                results.at[idx, param_key] = "NO VALUE FOUND"
                
                for pattern in PROTPARAM_PATTERNS[param_key]:
                    match = re.search(pattern, html, re.IGNORECASE | re.DOTALL)
                    if match:
                        try:
                            value_str = match.group(1).replace(',', '')
                            if param_key in ['mw', 'pi', 'gravy']:
                                value = float(value_str)
                            else:
                                value = int(value_str)
                            results.at[idx, param_key] = value
                            break
                        except (ValueError, IndexError):
                            continue
        
        if options.get('amino_acid', False):
            self._parse_amino_acids(results, idx, html, safe_mode)
    
    def _parse_amino_acids(self, results, idx, html, safe_mode):
        """Parse amino acid composition"""
        for aa_key in AMINO_ACID_COLUMNS.keys():
            if aa_key not in results.columns:
                results[aa_key] = "NO VALUE FOUND"
        
        for aa_key, pattern in AMINO_ACID_PATTERNS.items():
            if aa_key in results.columns and self.should_update(results, idx, aa_key, safe_mode):
                match = re.search(pattern, html, re.IGNORECASE)
                if match:
                    try:
                        count = int(match.group(1).strip())
                        percent = float(match.group(2).strip())
                        results.at[idx, aa_key] = f"{count}_{percent:.1f}%"
                    except (ValueError, IndexError):
                        results.at[idx, aa_key] = "0_0.0%"
                else:
                    results.at[idx, aa_key] = "0_0.0%"
        
        if 'atomic_comp' in results.columns and self.should_update(results, idx, 'atomic_comp', safe_mode):
            formula_match = re.search(ATOMIC_FORMULA_PATTERN, html, re.IGNORECASE)
            if formula_match:
                c, h, n, o, s = formula_match.groups()
                results.at[idx, 'atomic_comp'] = f"C{c}H{h}N{n}O{o}S{s}"
            else:
                results.at[idx, 'atomic_comp'] = "NO VALUE FOUND"
    
    def _set_no_value(self, results, idx, options, safe_mode):
        """Set ProtParam fields to NO VALUE FOUND"""
        protparam_fields = ['mw', 'pi', 'gravy', 'ext']
        self.set_no_value(results, idx, protparam_fields, safe_mode)
        
        if options.get('amino_acid', False):
            aa_fields = list(AMINO_ACID_COLUMNS.keys())
            for aa_key in aa_fields:
                if aa_key not in results.columns:
                    results[aa_key] = "NO VALUE FOUND"
                elif self.should_update(results, idx, aa_key, safe_mode):
                    results.at[idx, aa_key] = "NO VALUE FOUND"

class BLASTAnalyzer(BaseAnalyzer):
    """BLAST sequence similarity analyzer"""
    
    def __init__(self):
        super().__init__("BLAST")
    
    def analyze(self, results, options, progress_callback=None):
        """Run BLAST analysis"""
        self.logger.info("Starting BLAST analysis")
        
        safe_mode = options.get('safe_mode', True)
        to_process = self._get_processing_list(results, safe_mode)
        
        if not to_process:
            self.logger.info("All BLAST data complete")
            return
        
        self.logger.info(f"Processing {len(to_process)} sequences")
        
        for i, (idx, sequence, uniprot_id) in enumerate(to_process):
            if progress_callback:
                progress = 75 + (15 * (i + 1) / len(to_process))
                progress_callback(progress, f"BLAST ({i+1}/{len(to_process)})", f"Searching {uniprot_id}")
            
            clean_seq = re.sub(r'[^ACDEFGHIKLMNPQRSTVWY]', '', sequence.upper())
            self._set_no_value(results, idx, safe_mode)
            
            try:
                rid = self._submit_blast(clean_seq)
                if rid:
                    blast_results = self._wait_for_blast(rid)
                    if blast_results:
                        self._process_results(results, idx, blast_results, safe_mode)
            except Exception as e:
                self.logger.error(f"BLAST error for {uniprot_id}: {e}")
            
            if i < len(to_process) - 1:
                time.sleep(10)
    
    def _get_processing_list(self, results, safe_mode):
        """Get sequences that need BLAST processing"""
        to_process = []
        blast_fields = ['similar', 'identity', 'evalue', 'align']
        
        for idx, row in results.iterrows():
            sequence = row.get('sequence', '')
            uniprot_id = row.get('UniProt_ID', '')
            
            if sequence and sequence != "NO VALUE FOUND" and len(sequence) >= 20:
                if any(self.should_update(results, idx, field, safe_mode) for field in blast_fields):
                    to_process.append((idx, sequence, uniprot_id))
            else:
                self._set_no_value(results, idx, safe_mode)
        
        return to_process
    
    def _submit_blast(self, sequence):
        """Submit BLAST search to NCBI"""
        params = {
            'CMD': 'Put',
            'PROGRAM': 'blastp',
            'DATABASE': 'nr',
            'QUERY': sequence,
            'EXPECT': '10',
            'FORMAT_TYPE': 'XML',
            'EMAIL': 'protmerge@example.com',
            'TOOL': 'ProtMerge'
        }
        
        response = self.make_request(BLAST_URL, method='POST', data=params)
        
        if response and response.status_code == 200:
            rid_match = re.search(r'RID = ([A-Z0-9]+)', response.text)
            if rid_match:
                return rid_match.group(1)
        return None
    
    def _wait_for_blast(self, rid, max_wait=300):
        """Wait for BLAST results"""
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            try:
                status_params = {'CMD': 'Get', 'FORMAT_OBJECT': 'SearchInfo', 'RID': rid}
                status_response = self.make_request(BLAST_URL, params=status_params)
                
                if status_response and status_response.status_code == 200:
                    if 'Status=READY' in status_response.text:
                        return self._fetch_results(rid)
                    elif 'Status=WAITING' in status_response.text:
                        time.sleep(60)
                    elif 'Status=FAILED' in status_response.text:
                        break
                    else:
                        time.sleep(30)
                else:
                    break
            except Exception:
                break
        
        return None
    
    def _fetch_results(self, rid):
        """Fetch BLAST XML results"""
        result_params = {'CMD': 'Get', 'FORMAT_TYPE': 'XML', 'RID': rid}
        response = self.make_request(BLAST_URL, params=result_params)
        
        if response and response.status_code == 200:
            return self._parse_xml(response.text)
        return None
    
    def _parse_xml(self, xml_content):
        """Parse BLAST XML results"""
        try:
            root = ET.fromstring(xml_content)
            
            for hit in root.findall('.//Hit'):
                hit_def = hit.find('Hit_def')
                hsp = hit.find('.//Hsp')
                
                if hit_def is not None and hsp is not None:
                    identity_elem = hsp.find('Hsp_identity')
                    align_len_elem = hsp.find('Hsp_align-len')
                    evalue_elem = hsp.find('Hsp_evalue')
                    
                    if all(elem is not None for elem in [identity_elem, align_len_elem, evalue_elem]):
                        identity = int(identity_elem.text)
                        align_len = int(align_len_elem.text)
                        evalue = float(evalue_elem.text)
                        
                        percent_identity = (identity / align_len) * 100
                        
                        if percent_identity < 95:
                            return {
                                'similar': hit_def.text,
                                'identity': round(percent_identity, 2),
                                'evalue': evalue,
                                'align': align_len
                            }
            return None
        except Exception:
            return None
    
    def _process_results(self, results, idx, blast_data, safe_mode):
        """Process and store BLAST results"""
        for key, value in blast_data.items():
            if self.should_update(results, idx, key, safe_mode):
                results.at[idx, key] = value
    
    def _set_no_value(self, results, idx, safe_mode):
        """Set BLAST fields to NO VALUE FOUND"""
        blast_fields = ['similar', 'identity', 'evalue', 'align']
        self.set_no_value(results, idx, blast_fields, safe_mode)

class PDBAnalyzer(BaseAnalyzer):
    """PDB structure analyzer"""
    
    def __init__(self):
        super().__init__("PDB")
    
    def analyze(self, results, options, progress_callback=None):
        """Run PDB structure analysis"""
        self.logger.info("Starting PDB structure analysis")
        
        safe_mode = options.get('safe_mode', True)
        
        pdb_fields = ['structure_count', 'best_resolution', 'structure_methods', 'complex_info',
                     'pdb_ids', 'best_structure', 'ligand_info', 'structure_quality']
        
        for field in pdb_fields:
            if field not in results.columns:
                results[field] = "NO VALUE FOUND"
        
        to_process = self._get_processing_list(results, safe_mode)
        
        if not to_process:
            self.logger.info("All PDB data complete")
            return
        
        self.logger.info(f"Processing {len(to_process)} proteins")
        
        for i, (idx, uniprot_id) in enumerate(to_process):
            if progress_callback:
                progress = 85 + (10 * (i + 1) / len(to_process))
                progress_callback(progress, f"PDB ({i+1}/{len(to_process)})", f"Searching {uniprot_id}")
            
            try:
                pdb_entries = self._search_structures(uniprot_id)
                
                if pdb_entries:
                    pdb_ids = [entry['identifier'] for entry in pdb_entries]
                    structure_details = self._get_structure_details(pdb_ids[:10])
                    summary = self._summarize_structures(pdb_ids, structure_details)
                    self._process_results(results, idx, summary, safe_mode)
                else:
                    self._set_no_structures(results, idx, safe_mode)
                    
            except Exception as e:
                self.logger.error(f"PDB error for {uniprot_id}: {e}")
                self._set_no_value(results, idx, safe_mode)
            
            time.sleep(0.5)
    
    def _get_processing_list(self, results, safe_mode):
        """Get proteins that need PDB processing"""
        to_process = []
        pdb_fields = ['structure_count', 'best_resolution', 'structure_methods', 'complex_info',
                     'pdb_ids', 'best_structure', 'ligand_info', 'structure_quality']
        
        for idx, row in results.iterrows():
            uniprot_id = row.get('UniProt_ID', '')
            if uniprot_id and any(self.should_update(results, idx, field, safe_mode) for field in pdb_fields):
                to_process.append((idx, uniprot_id))
        
        return to_process
    
    def _search_structures(self, uniprot_id):
        """Search PDB structures using UniProt ID"""
        search_query = {
            "query": {
                "type": "terminal",
                "service": "text",
                "parameters": {
                    "attribute": "rcsb_polymer_entity_container_identifiers.reference_sequence_identifiers.database_accession",
                    "operator": "exact_match",
                    "value": uniprot_id
                }
            },
            "return_type": "entry",
            "request_options": {"return_all_hits": True}
        }
        
        try:
            response = self.make_request(PDB_SEARCH_URL, method='POST', json=search_query)
            if response and response.status_code == 200:
                data = response.json()
                return data.get("result_set", [])
            return []
        except Exception:
            return []
    
    def _get_structure_details(self, pdb_ids):
        """Get detailed information for PDB structures"""
        structure_data = {}
        
        for pdb_id in pdb_ids[:10]:
            try:
                url = f"{PDB_DATA_URL}/entry/{pdb_id}"
                response = self.make_request(url)
                
                if response and response.status_code == 200:
                    data = response.json()
                    structure_data[pdb_id] = self._extract_structure_info(data)
            except Exception:
                continue
        
        return structure_data
    
    def _extract_structure_info(self, data):
        """Extract key information from PDB entry data"""
        info = {
            'resolution': 'N/A',
            'method': 'Unknown',
            'title': 'N/A',
            'ligands': []
        }
        
        try:
            if 'rcsb_entry_info' in data:
                entry_info = data['rcsb_entry_info']
                resolution = entry_info.get('resolution_combined')
                if resolution and len(resolution) > 0:
                    info['resolution'] = resolution[0]
            
            if 'exptl' in data:
                methods = []
                for exp in data['exptl']:
                    method = exp.get('method', 'Unknown')
                    if method and method not in methods:
                        methods.append(method)
                info['method'] = '; '.join(methods) if methods else 'Unknown'
            
            if 'struct' in data:
                info['title'] = data['struct'].get('title', 'N/A')
            
            title = info['title'].lower()
            ligand_keywords = {
                'dna': 'DNA', 'rna': 'RNA', 'atp': 'ATP', 'inhibitor': 'Inhibitor',
                'drug': 'Drug compound', 'substrate': 'Substrate', 'cofactor': 'Cofactor'
            }
            
            for keyword, ligand_name in ligand_keywords.items():
                if keyword in title:
                    info['ligands'].append(ligand_name)
            
        except Exception:
            pass
        
        return info
    
    def _summarize_structures(self, pdb_ids, structure_details):
        """Summarize PDB structure information"""
        if not structure_details:
            return {
                'structure_count': 0,
                'best_resolution': 'N/A',
                'methods': 'No structures available',
                'complex_info': 'N/A',
                'pdb_ids': 'None',
                'best_structure': 'N/A',
                'ligand_info': 'N/A',
                'structure_quality': 'N/A'
            }
        
        summary = {
            'structure_count': len(pdb_ids),
            'best_resolution': 'N/A',
            'methods': set(),
            'complex_info': set(),
            'pdb_ids': '; '.join(pdb_ids[:15]),
            'best_structure': pdb_ids[0] if pdb_ids else 'N/A',
            'ligand_info': set(),
            'structure_quality': 'Unknown'
        }
        
        best_resolution = float('inf')
        
        for pdb_id, details in structure_details.items():
            if details['method'] != 'Unknown':
                summary['methods'].add(details['method'])
            
            try:
                resolution = float(details['resolution'])
                if resolution < best_resolution:
                    best_resolution = resolution
                    summary['best_structure'] = pdb_id
            except (ValueError, TypeError):
                pass
            
            title = details['title'].lower()
            if 'complex' in title or 'bound' in title:
                if any(word in title for word in ['dna', 'rna']):
                    summary['complex_info'].add('Nucleic Acid Complex')
                elif any(word in title for word in ['drug', 'inhibitor']):
                    summary['complex_info'].add('Drug Complex')
                else:
                    summary['complex_info'].add('Ligand Complex')
            
            summary['ligand_info'].update(details['ligands'])
        
        summary['best_resolution'] = f"{best_resolution:.2f}Å" if best_resolution != float('inf') else 'N/A'
        summary['methods'] = '; '.join(sorted(summary['methods'])) if summary['methods'] else 'Unknown'
        summary['complex_info'] = '; '.join(sorted(summary['complex_info'])) if summary['complex_info'] else 'Monomer'
        summary['ligand_info'] = '; '.join(sorted(list(summary['ligand_info']))[:5]) if summary['ligand_info'] else 'None'
        summary['structure_quality'] = self._assess_quality(best_resolution)
        
        return summary
    
    def _assess_quality(self, resolution):
        """Assess structure quality based on resolution"""
        if resolution == float('inf'):
            return 'Unknown'
        elif resolution <= 1.5:
            return 'Excellent'
        elif resolution <= 2.0:
            return 'Very Good'
        elif resolution <= 2.5:
            return 'Good'
        elif resolution <= 3.5:
            return 'Moderate'
        else:
            return 'Low'
    
    def _process_results(self, results, idx, summary, safe_mode):
        """Process and store PDB results"""
        pdb_field_mapping = {
            'structure_count': 'structure_count',
            'best_resolution': 'best_resolution',
            'structure_methods': 'methods',
            'complex_info': 'complex_info',
            'pdb_ids': 'pdb_ids',
            'best_structure': 'best_structure',
            'ligand_info': 'ligand_info',
            'structure_quality': 'structure_quality'
        }
        
        for result_key, summary_key in pdb_field_mapping.items():
            if self.should_update(results, idx, result_key, safe_mode):
                results.at[idx, result_key] = summary[summary_key]
    
    def _set_no_structures(self, results, idx, safe_mode):
        """Set specific message when no structures are found"""
        pdb_data = {
            'structure_count': 0,
            'best_resolution': 'No structures available',
            'structure_methods': 'No structures available',
            'complex_info': 'No structures available',
            'pdb_ids': 'None found',
            'best_structure': 'No structures available',
            'ligand_info': 'No structures available',
            'structure_quality': 'No structures available'
        }
        
        for field, value in pdb_data.items():
            if self.should_update(results, idx, field, safe_mode):
                results.at[idx, field] = value
    
    def _set_no_value(self, results, idx, safe_mode):
        """Set PDB fields to NO VALUE FOUND"""
        pdb_fields = ['structure_count', 'best_resolution', 'structure_methods', 'complex_info',
                     'pdb_ids', 'best_structure', 'ligand_info', 'structure_quality']
        self.set_no_value(results, idx, pdb_fields, safe_mode)