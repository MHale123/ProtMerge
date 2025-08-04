import requests
import time
import json
import logging
import pandas as pd
import xml.etree.ElementTree as ET
import re
from urllib.parse import quote, urlencode
from config import *

class HumanProteinAnalyzerManager:
    """Enhanced version with comprehensive COMPARTMENTS data extraction"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.ensembl_mapper = EnsemblGeneMapper()
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1'
        })
        
        # COMPARTMENTS confidence mapping
        self.confidence_map = {
            5: "★★★★★", 4: "★★★★☆", 3: "★★★☆☆", 
            2: "★★☆☆☆", 1: "★☆☆☆☆", 0: "☆☆☆☆☆"
        }
        
        self.logger.info("Enhanced HumanProteinAnalyzerManager initialized")
    
    def run_human_analysis(self, data, options, progress_callback=None):
        """Run human-specific analyses with enhanced COMPARTMENTS extraction"""
        results = data['results']
        
        # Initialize human protein columns
        self._initialize_human_columns(results)
        
        # Check if we have human-specific analyses to run
        has_human_analysis = options.get('compartments', False) or options.get('hpa', False)
        if not has_human_analysis:
            self.logger.info("No human-specific analyses requested")
            return results
        
        self.logger.info("Starting enhanced human-specific protein analysis")
        
        # Get gene names from UniProt_ID column (these are actually gene names at this point)
        gene_names = results['UniProt_ID'].tolist()
        total_genes = len(gene_names)
        
        self.logger.info(f"Processing {total_genes} genes for human-specific analysis")
        
        # Run enhanced COMPARTMENTS analysis if requested
        if options.get('compartments', False):
            self.logger.info("Starting enhanced COMPARTMENTS analysis...")
            self._run_compartments_analysis_enhanced(results, gene_names, progress_callback, total_genes)
        
        # Run HPA analysis if requested  
        if options.get('hpa', False):
            self.logger.info("Starting HPA analysis...")
            self._run_hpa_analysis_fixed(results, gene_names, progress_callback, total_genes)
        
        # Log final statistics
        self._log_analysis_statistics(results, options)
        
        self.logger.info("Enhanced human-specific analysis completed")
        return results
    
    def _run_compartments_analysis_enhanced(self, results, gene_names, progress_callback, total_genes):
        """Enhanced COMPARTMENTS analysis with comprehensive data extraction"""
        self.logger.info("Running enhanced COMPARTMENTS analysis with multiple data sources")
        
        for i, gene_name in enumerate(gene_names):
            if progress_callback:
                progress = 5 + (15 * (i + 1) / total_genes)
                progress_callback(progress, f"COMPARTMENTS ({i+1}/{total_genes})", f"Analyzing {gene_name}")
            
            try:
                clean_gene = str(gene_name).strip()
                self.logger.debug(f"Enhanced COMPARTMENTS: Processing {clean_gene}")
                
                # Try multiple approaches to get comprehensive COMPARTMENTS data
                compartments_data = self._get_compartments_comprehensive(clean_gene)
                self._process_compartments_data_enhanced(results, i, compartments_data, clean_gene)
                
            except Exception as e:
                self.logger.error(f"Enhanced COMPARTMENTS error for {gene_name}: {e}")
                self._set_compartments_no_value(results, i)
            
            time.sleep(RATE_LIMITS.get('compartments', 0.3))
    
    def _get_compartments_comprehensive(self, gene_name):
        """Get comprehensive COMPARTMENTS data using multiple strategies"""
        all_locations = []
        
        # Strategy 1: Direct COMPARTMENTS API (if available)
        compartments_direct = self._get_compartments_direct_api(gene_name)
        if compartments_direct:
            all_locations.extend(compartments_direct)
            self.logger.debug(f"COMPARTMENTS direct API: {len(compartments_direct)} locations for {gene_name}")
        
        # Strategy 2: Jensen Lab COMPARTMENTS via STRING/UniProt
        compartments_jensen = self._get_compartments_jensen_lab(gene_name)
        if compartments_jensen:
            all_locations.extend(compartments_jensen)
            self.logger.debug(f"Jensen Lab COMPARTMENTS: {len(compartments_jensen)} locations for {gene_name}")
        
        # Strategy 3: Enhanced UniProt subcellular location
        uniprot_data = self._get_uniprot_subcellular_enhanced(gene_name)
        if uniprot_data:
            all_locations.extend(uniprot_data)
            self.logger.debug(f"UniProt enhanced: {len(uniprot_data)} locations for {gene_name}")
        
        # Strategy 4: GO Cellular Component (enhanced)
        go_data = self._get_go_cellular_component_enhanced(gene_name)
        if go_data:
            all_locations.extend(go_data)
            self.logger.debug(f"GO enhanced: {len(go_data)} locations for {gene_name}")
        
        # Strategy 5: Literature mining (PubMed abstracts)
        literature_data = self._get_literature_compartments(gene_name)
        if literature_data:
            all_locations.extend(literature_data)
            self.logger.debug(f"Literature mining: {len(literature_data)} locations for {gene_name}")
        
        # Consolidate and rank data
        if all_locations:
            consolidated = self._consolidate_compartments_data(all_locations, gene_name)
            self.logger.info(f"COMPARTMENTS comprehensive success for {gene_name}: {len(consolidated)} unique locations")
            return consolidated
        
        self.logger.debug(f"No COMPARTMENTS data found for {gene_name}")
        return None
    
    def _get_compartments_direct_api(self, gene_name):
        """Try to access COMPARTMENTS database directly"""
        try:
            # Try the COMPARTMENTS API endpoint (if available)
            base_urls = [
                "https://compartments.jensenlab.org/api",
                "https://diseases.jensenlab.org/api",
                "http://compartments.jensenlab.org/api"
            ]
            
            for base_url in base_urls:
                try:
                    url = f"{base_url}/entities"
                    params = {
                        'query': gene_name,
                        'format': 'json',
                        'limit': 20
                    }
                    
                    response = self.session.get(url, params=params, timeout=15)
                    
                    if response.status_code == 200:
                        data = response.json()
                        locations = []
                        
                        for item in data.get('results', []):
                            location_name = item.get('name', '')
                            confidence = item.get('confidence', 0)
                            evidence = item.get('evidence', 'Unknown')
                            source = item.get('source', 'COMPARTMENTS')
                            
                            if location_name:
                                locations.append({
                                    'name': location_name,
                                    'source': source,
                                    'evidence': evidence,
                                    'confidence': self._convert_confidence_to_stars(confidence),
                                    'score': confidence
                                })
                        
                        if locations:
                            return locations
                
                except Exception as e:
                    self.logger.debug(f"COMPARTMENTS API {base_url} failed for {gene_name}: {e}")
                    continue
            
            return None
            
        except Exception as e:
            self.logger.debug(f"COMPARTMENTS direct API failed for {gene_name}: {e}")
            return None
    
    def _get_compartments_jensen_lab(self, gene_name):
        """Get data from Jensen Lab COMPARTMENTS via alternative endpoints"""
        try:
            # Try STRING database integration (Jensen Lab maintains STRING)
            string_id = self._get_string_protein_id(gene_name)
            if string_id:
                url = "https://string-db.org/api/json/functional_annotation"
                params = {
                    'identifiers': string_id,
                    'species': 9606,
                    'caller_identity': 'ProtMerge'
                }
                
                response = self.session.get(url, params=params, timeout=20)
                
                if response.status_code == 200:
                    data = response.json()
                    locations = []
                    
                    for item in data:
                        category = item.get('category', '').lower()
                        description = item.get('description', '')
                        
                        # Look for compartment/localization annotations
                        if any(keyword in category for keyword in ['component', 'localization', 'compartment']):
                            # Extract location information
                            location_info = self._extract_location_from_description(description)
                            if location_info:
                                locations.append({
                                    'name': location_info['name'],
                                    'source': 'STRING/COMPARTMENTS',
                                    'evidence': 'IEA',  # Inferred from Electronic Annotation
                                    'confidence': '★★★☆☆',  # Medium confidence for STRING
                                    'score': 3
                                })
                    
                    return locations
            
            return None
            
        except Exception as e:
            self.logger.debug(f"Jensen Lab COMPARTMENTS failed for {gene_name}: {e}")
            return None
    
    def _get_string_protein_id(self, gene_name):
        """Get STRING protein ID for gene"""
        try:
            url = "https://string-db.org/api/json/get_string_ids"
            params = {
                'identifiers': gene_name,
                'species': 9606,
                'caller_identity': 'ProtMerge'
            }
            
            response = self.session.get(url, params=params, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                if data and len(data) > 0:
                    return data[0].get('stringId')
            
            return None
            
        except Exception:
            return None
    
    def _get_uniprot_subcellular_enhanced(self, gene_name):
        """Enhanced UniProt subcellular location extraction"""
        try:
            url = "https://rest.uniprot.org/uniprotkb/search"
            params = {
                'query': f'gene:{gene_name} AND organism_id:9606',
                'format': 'json',
                'fields': 'accession,gene_names,cc_subcellular_location,ft_topo_dom,ft_transmem,ft_signal',
                'size': 3  # Get top 3 results
            }
            
            response = self.session.get(url, params=params, timeout=15)
            
            if response.status_code == 200:
                data = response.json()
                results = data.get('results', [])
                
                locations = []
                
                for protein in results:
                    # Extract subcellular location comments
                    comments = protein.get('comments', [])
                    for comment in comments:
                        if comment.get('commentType') == 'SUBCELLULAR LOCATION':
                            subcellular_locations = comment.get('subcellularLocations', [])
                            for loc in subcellular_locations:
                                location_name = loc.get('location', {}).get('value', '')
                                if location_name:
                                    locations.append({
                                        'name': location_name,
                                        'source': 'UniProt',
                                        'evidence': 'IEA',
                                        'confidence': '★★★★☆',  # High confidence for UniProt
                                        'score': 4
                                    })
                    
                    # Extract features (transmembrane, signal peptides, etc.)
                    features = protein.get('features', [])
                    for feature in features:
                        feature_type = feature.get('type', '')
                        if feature_type in ['TRANSMEM', 'SIGNAL', 'TOPO_DOM']:
                            if feature_type == 'TRANSMEM':
                                locations.append({
                                    'name': 'Membrane',
                                    'source': 'UniProt',
                                    'evidence': 'IEA',
                                    'confidence': '★★★★☆',
                                    'score': 4
                                })
                            elif feature_type == 'SIGNAL':
                                locations.append({
                                    'name': 'Secreted',
                                    'source': 'UniProt',
                                    'evidence': 'IEA',
                                    'confidence': '★★★★☆',
                                    'score': 4
                                })
                
                return locations
            
            return None
            
        except Exception as e:
            self.logger.debug(f"UniProt enhanced query failed for {gene_name}: {e}")
            return None
    
    def _get_go_cellular_component_enhanced(self, gene_name):
        """Enhanced GO Cellular Component extraction"""
        try:
            # First get UniProt ID(s)
            url = "https://rest.uniprot.org/uniprotkb/search"
            params = {
                'query': f'gene:{gene_name} AND organism_id:9606',
                'format': 'json',
                'fields': 'accession',
                'size': 3
            }
            
            response = self.session.get(url, params=params, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                results = data.get('results', [])
                
                all_locations = []
                
                for protein in results:
                    uniprot_id = protein.get('primaryAccession')
                    
                    if uniprot_id:
                        # Query QuickGO for detailed GO annotations
                        go_locations = self._query_quickgo_detailed(uniprot_id)
                        if go_locations:
                            all_locations.extend(go_locations)
                
                return all_locations
            
            return None
            
        except Exception as e:
            self.logger.debug(f"GO enhanced query failed for {gene_name}: {e}")
            return None
    
    def _query_quickgo_detailed(self, uniprot_id):
        """Detailed QuickGO query with evidence extraction"""
        try:
            url = "https://www.ebi.ac.uk/QuickGO/services/annotation/search"
            params = {
                'geneProductId': uniprot_id,
                'aspect': 'cellular_component',
                'limit': 30,
                'includeFields': 'goName,evidenceCode,reference,qualifier'
            }
            
            response = self.session.get(url, params=params, timeout=20)
            
            if response.status_code == 200:
                data = response.json()
                annotations = data.get('results', [])
                
                locations = []
                
                for annotation in annotations:
                    go_name = annotation.get('goName', '')
                    evidence_code = annotation.get('evidenceCode', 'IEA')
                    qualifier = annotation.get('qualifier', '')
                    
                    if go_name and 'NOT' not in qualifier:
                        # Clean up GO term names
                        clean_name = go_name.replace(' (sensu ', ' ').replace(')', '')
                        
                        # Map evidence codes to confidence
                        confidence = self._map_evidence_to_confidence(evidence_code)
                        
                        locations.append({
                            'name': clean_name,
                            'source': 'GO/QuickGO',
                            'evidence': evidence_code,
                            'confidence': confidence,
                            'score': self._confidence_to_score(confidence)
                        })
                
                return locations
            
            return None
            
        except Exception as e:
            self.logger.debug(f"QuickGO detailed query failed for {uniprot_id}: {e}")
            return None
    
    def _get_literature_compartments(self, gene_name):
        """Extract compartment information from literature (PubMed abstracts)"""
        try:
            # Search PubMed for recent papers about the gene and subcellular localization
            search_terms = f"{gene_name} AND (subcellular localization OR cellular compartment OR intracellular)"
            
            # Use NCBI E-utilities
            search_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
            search_params = {
                'db': 'pubmed',
                'term': search_terms,
                'retmax': 10,
                'retmode': 'json',
                'sort': 'relevance'
            }
            
            search_response = self.session.get(search_url, params=search_params, timeout=15)
            
            if search_response.status_code == 200:
                search_data = search_response.json()
                pmids = search_data.get('esearchresult', {}).get('idlist', [])
                
                if pmids:
                    # Fetch abstracts
                    fetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
                    fetch_params = {
                        'db': 'pubmed',
                        'id': ','.join(pmids[:5]),  # Limit to 5 abstracts
                        'retmode': 'xml'
                    }
                    
                    fetch_response = self.session.get(fetch_url, params=fetch_params, timeout=20)
                    
                    if fetch_response.status_code == 200:
                        # Parse abstracts for location keywords
                        locations = self._extract_locations_from_abstracts(fetch_response.text, gene_name)
                        return locations
            
            return None
            
        except Exception as e:
            self.logger.debug(f"Literature mining failed for {gene_name}: {e}")
            return None
    
    def _extract_locations_from_abstracts(self, xml_text, gene_name):
        """Extract location information from PubMed abstracts"""
        try:
            locations = []
            
            # Define location keywords and their standardized names
            location_keywords = {
                'nucleus': 'Nucleus',
                'nuclear': 'Nucleus',
                'cytoplasm': 'Cytoplasm',
                'cytoplasmic': 'Cytoplasm',
                'mitochondria': 'Mitochondrion',
                'mitochondrial': 'Mitochondrion',
                'endoplasmic reticulum': 'Endoplasmic reticulum',
                'er': 'Endoplasmic reticulum',
                'golgi': 'Golgi apparatus',
                'membrane': 'Membrane',
                'plasma membrane': 'Cell membrane',
                'cell membrane': 'Cell membrane',
                'ribosome': 'Ribosome',
                'ribosomal': 'Ribosome',
                'lysosome': 'Lysosome',
                'lysosomal': 'Lysosome',
                'peroxisome': 'Peroxisome',
                'secreted': 'Secreted',
                'extracellular': 'Extracellular region'
            }
            
            # Convert to lowercase for searching
            text_lower = xml_text.lower()
            
            found_locations = set()
            for keyword, standard_name in location_keywords.items():
                if keyword in text_lower:
                    found_locations.add(standard_name)
            
            # Convert to location objects
            for location in found_locations:
                locations.append({
                    'name': location,
                    'source': 'Literature',
                    'evidence': 'TAS',  # Traceable Author Statement
                    'confidence': '★★☆☆☆',  # Lower confidence for text mining
                    'score': 2
                })
            
            return locations
            
        except Exception as e:
            self.logger.debug(f"Abstract parsing failed: {e}")
            return None
    
    def _consolidate_compartments_data(self, all_locations, gene_name):
        """Consolidate and rank COMPARTMENTS data from multiple sources with practical output format"""
        try:
            # Group by location name
            location_groups = {}
            
            for loc in all_locations:
                name = loc['name']
                if name not in location_groups:
                    location_groups[name] = []
                location_groups[name].append(loc)
            
            # Consolidate each location
            consolidated = []
            
            for location_name, group in location_groups.items():
                # Find the highest confidence source
                best_item = max(group, key=lambda x: x['score'])
                
                # Collect all sources and evidence types
                sources = list(set(item['source'] for item in group))
                evidences = list(set(item['evidence'] for item in group))
                
                consolidated_item = {
                    'name': location_name,
                    'confidence_score': best_item['score'],  # Numerical for sorting
                    'sources': sources,
                    'evidences': evidences,
                    'source_count': len(sources)
                }
                
                consolidated.append(consolidated_item)
            
            # Sort by confidence score (highest first), then by source count
            consolidated.sort(key=lambda x: (x['confidence_score'], x['source_count']), reverse=True)
            
            return consolidated[:15]  # Return top 15 locations
            
        except Exception as e:
            self.logger.error(f"Error consolidating COMPARTMENTS data for {gene_name}: {e}")
            return None
    
    def _process_compartments_data_enhanced(self, results, idx, data, gene_name):
        """Process COMPARTMENTS data with practical, sortable format"""
        if not data:
            self._set_compartments_no_value(results, idx)
            return
        
        try:
            # Get primary (highest confidence) location
            primary_location = data[0]['name']
            primary_confidence = data[0]['confidence_score']
            
            # Prepare all locations data
            all_locations = [item['name'] for item in data]
            all_confidence_scores = [str(item['confidence_score']) for item in data]
            all_sources = []
            all_evidences = []
            
            for item in data:
                all_sources.extend(item['sources'])
                all_evidences.extend(item['evidences'])
            
            # Remove duplicates while preserving order
            unique_sources = list(dict.fromkeys(all_sources))
            unique_evidences = list(dict.fromkeys(all_evidences))
            
            # Store in results with practical format
            results.at[idx, 'compartments_primary_location'] = primary_location
            results.at[idx, 'compartments_primary_confidence'] = primary_confidence
            results.at[idx, 'compartments_all_locations'] = " | ".join(all_locations)
            results.at[idx, 'compartments_confidence_scores'] = " | ".join(all_confidence_scores)
            results.at[idx, 'compartments_evidence_types'] = " | ".join(unique_evidences)
            results.at[idx, 'compartments_data_sources'] = " | ".join(unique_sources)
            
            self.logger.info(f"COMPARTMENTS success for {gene_name}: Primary={primary_location} (confidence={primary_confidence}), Total={len(data)} locations")
            
        except Exception as e:
            self.logger.error(f"Error processing COMPARTMENTS data for {gene_name}: {e}")
            self._set_compartments_no_value(results, idx)
    
    # Helper methods
    def _convert_confidence_to_stars(self, confidence_score):
        """Convert numerical confidence to star rating"""
        if isinstance(confidence_score, str):
            try:
                confidence_score = float(confidence_score)
            except:
                return "★★★☆☆"  # Default medium confidence
        
        if confidence_score >= 4.5:
            return "★★★★★"
        elif confidence_score >= 3.5:
            return "★★★★☆"
        elif confidence_score >= 2.5:
            return "★★★☆☆"
        elif confidence_score >= 1.5:
            return "★★☆☆☆"
        elif confidence_score >= 0.5:
            return "★☆☆☆☆"
        else:
            return "☆☆☆☆☆"
    
    def _map_evidence_to_confidence(self, evidence_code):
        """Map GO evidence codes to confidence levels"""
        high_confidence = ['EXP', 'IDA', 'IPI', 'IMP', 'IGI', 'IEP']
        medium_confidence = ['ISS', 'ISO', 'ISA', 'ISM', 'IGC', 'IBA', 'IBD', 'IKR', 'IRD']
        low_confidence = ['IEA', 'TAS', 'NAS', 'IC', 'ND']
        
        if evidence_code in high_confidence:
            return "★★★★★"
        elif evidence_code in medium_confidence:
            return "★★★☆☆"
        elif evidence_code in low_confidence:
            return "★★☆☆☆"
        else:
            return "★★★☆☆"  # Default
    
    def _confidence_to_score(self, confidence_stars):
        """Convert star confidence to numerical score"""
        return confidence_stars.count("★")
    
    def _extract_location_from_description(self, description):
        """Extract location information from text description"""
        description_lower = description.lower()
        
        location_mapping = {
            'nucleus': 'Nucleus',
            'cytoplasm': 'Cytoplasm',
            'mitochondrion': 'Mitochondrion',
            'membrane': 'Membrane',
            'golgi': 'Golgi apparatus',
            'ribosome': 'Ribosome',
            'lysosome': 'Lysosome',
            'peroxisome': 'Peroxisome',
            'endoplasmic reticulum': 'Endoplasmic reticulum'
        }
        
        for keyword, standard_name in location_mapping.items():
            if keyword in description_lower:
                return {'name': standard_name}
        
        return None
    
    # Keep existing methods for HPA analysis and initialization
    def _initialize_human_columns(self, results):
        """Initialize columns for human protein data"""
        for key in HUMAN_PROTEIN_COLUMNS.keys():
            if key not in results.columns:
                results[key] = "NO VALUE FOUND"
        
        # Add new detailed column for enhanced COMPARTMENTS
        if 'compartments_detailed' not in results.columns:
            results['compartments_detailed'] = "NO VALUE FOUND"
    
    def _run_hpa_analysis_fixed(self, results, gene_names, progress_callback, total_genes):
        """Enhanced HPA analysis with multiple data retrieval strategies"""
        self.logger.info("Running enhanced HPA analysis with multiple approaches")
        
        for i, gene_name in enumerate(gene_names):
            if progress_callback:
                progress = 20 + (15 * (i + 1) / total_genes)
                progress_callback(progress, f"HPA ({i+1}/{total_genes})", f"Analyzing {gene_name}")
            
            try:
                clean_gene = str(gene_name).strip()
                self.logger.debug(f"HPA: Processing {clean_gene}")
                
                # Try multiple HPA data retrieval strategies
                hpa_data = self._get_hpa_comprehensive(clean_gene)
                self._process_hpa_data_enhanced(results, i, hpa_data, clean_gene)
                
            except Exception as e:
                self.logger.error(f"HPA error for {gene_name}: {e}")
                self._set_hpa_no_value(results, i)
            
            time.sleep(RATE_LIMITS.get('hpa', 1.0))
    
    def _get_hpa_comprehensive(self, gene_name):
        """Comprehensive HPA data retrieval using multiple strategies"""
        # Strategy 1: Direct HPA API/XML (if available)
        ensembl_id = self.ensembl_mapper.get_ensembl_id(gene_name)
        if ensembl_id:
            hpa_data = self._get_hpa_xml_enhanced(ensembl_id, gene_name)
            if hpa_data and self._is_valid_hpa_data(hpa_data):
                self.logger.debug(f"HPA XML success for {gene_name}")
                return hpa_data
        
        # Strategy 2: HPA web scraping (careful parsing)
        hpa_data = self._get_hpa_web_safe(gene_name)
        if hpa_data and self._is_valid_hpa_data(hpa_data):
            self.logger.debug(f"HPA web scraping success for {gene_name}")
            return hpa_data
        
        # Strategy 3: Expression Atlas as HPA alternative
        atlas_data = self._get_expression_atlas_enhanced(gene_name)
        if atlas_data and self._is_valid_hpa_data(atlas_data):
            self.logger.debug(f"Expression Atlas success for {gene_name}")
            return atlas_data
        
        # Strategy 4: GTEx (Genotype-Tissue Expression) data
        gtex_data = self._get_gtex_data(gene_name)
        if gtex_data and self._is_valid_hpa_data(gtex_data):
            self.logger.debug(f"GTEx success for {gene_name}")
            return gtex_data
        
        # Strategy 5: UniProt tissue specificity (fallback)
        uniprot_tissue = self._get_uniprot_tissue_enhanced(gene_name)
        if uniprot_tissue and self._is_valid_hpa_data(uniprot_tissue):
            self.logger.debug(f"UniProt tissue success for {gene_name}")
            return uniprot_tissue
        
        self.logger.debug(f"No reliable HPA data found for {gene_name}")
        return None
    
    def _get_hpa_xml_enhanced(self, ensembl_id, gene_name):
        """Enhanced HPA XML parsing with better subcellular location extraction"""
        try:
            url = f"https://www.proteinatlas.org/{ensembl_id}.xml"
        
            response = self.session.get(url, timeout=30)
        
            if response.status_code == 200:
                try:
                    root = ET.fromstring(response.content)
                
                    tissues = set()
                    locations = set()
                    expression_levels = []
                    antibody_info = []
                
                    # FIXED: Enhanced subcellular location extraction with multiple approaches
                
                    # Approach 1: Look for subcellular location elements
                    for loc_elem in root.findall('.//subcellularLocation'):
                        # Try different ways to get location name
                        loc_name = None
                    
                        # Try as attribute
                        if loc_elem.get('name'):
                            loc_name = loc_elem.get('name')
                        # Try as text content
                        elif loc_elem.text:
                            loc_name = loc_elem.text.strip()
                        # Try nested elements
                        else:
                            for child in loc_elem:
                                if child.text and child.text.strip():
                                    loc_name = child.text.strip()
                                    break
                    
                        if loc_name and self._is_valid_location_name(loc_name):
                            locations.add(loc_name.strip())
                            self.logger.debug(f"Found subcellular location: {loc_name}")
                
                    # Approach 2: Look for location in different XML structures
                    location_xpath_patterns = [
                        './/location',
                        './/subcellular',
                        './/cellularComponent',
                        './/localization',
                        './/compartment'
                    ]
                
                    for xpath in location_xpath_patterns:
                        for elem in root.findall(xpath):
                            loc_name = self._extract_location_name(elem)
                            if loc_name and self._is_valid_location_name(loc_name):
                                locations.add(loc_name)
                                self.logger.debug(f"Found location via {xpath}: {loc_name}")
                
                    # Approach 3: Look in immunofluorescence data
                    for if_elem in root.findall('.//immunofluorescence'):
                        location_attr = if_elem.get('location', '')
                        if location_attr and self._is_valid_location_name(location_attr):
                            locations.add(location_attr)
                            self.logger.debug(f"Found IF location: {location_attr}")
                    
                        # Check nested location elements in IF data
                        for loc_child in if_elem.findall('.//location'):
                            loc_name = self._extract_location_name(loc_child)
                            if loc_name and self._is_valid_location_name(loc_name):
                                locations.add(loc_name)
                
                    # Approach 4: Look in antibody staining data
                    for antibody_elem in root.findall('.//antibody'):
                        for staining in antibody_elem.findall('.//staining'):
                            location_attr = staining.get('location', '')
                            if location_attr and self._is_valid_location_name(location_attr):
                                locations.add(location_attr)
                        
                            # Check for location in staining text
                            if staining.text:
                                parsed_locations = self._parse_location_from_text(staining.text)
                                locations.update(parsed_locations)
                
                    # Approach 5: Look for GO cellular component terms
                    for go_elem in root.findall('.//go'):
                        go_term = go_elem.get('term', '')
                        go_aspect = go_elem.get('aspect', '')
                    
                        if go_aspect.lower() == 'cellular_component' and go_term:
                            # Clean up GO term
                            clean_term = go_term.replace('GO:', '').strip()
                            if self._is_valid_location_name(clean_term):
                                locations.add(clean_term)
                                self.logger.debug(f"Found GO cellular component: {clean_term}")
                
                    # Extract tissue expression data (existing code)
                    for tissue_elem in root.findall('.//tissue'):
                        tissue_name = tissue_elem.get('name', tissue_elem.text)
                        level = tissue_elem.get('level', 'detected')
                    
                        if tissue_name and self._is_valid_tissue_name(tissue_name):
                            tissues.add(tissue_name.strip())
                            if level and level != 'not detected':
                                expression_levels.append(f"{tissue_name.strip()}:{level}")
                
                    # Extract expression data from different XML structures
                    for expr_elem in root.findall('.//expression'):
                        tissue = expr_elem.get('tissue', '')
                        level = expr_elem.get('level', '')
                    
                        if tissue and self._is_valid_tissue_name(tissue):
                            tissues.add(tissue)
                            if level != 'not detected':
                                expression_levels.append(f"{tissue}:{level}")
                
                    # Extract antibody reliability information
                    for antibody_elem in root.findall('.//antibody'):
                        reliability = antibody_elem.get('reliability', '')
                        if reliability:
                            antibody_info.append(reliability)
                
                    # Log what we found
                    if locations:
                        self.logger.info(f"HPA XML found {len(locations)} subcellular locations for {gene_name}: {list(locations)}")
                    else:
                        self.logger.warning(f"HPA XML found no subcellular locations for {gene_name}")
                
                    if tissues or locations or expression_levels:
                        return {
                            'gene_name': gene_name,
                            'tissues': list(tissues)[:20],
                            'subcellular_locations': list(locations)[:10],  # This should now have data
                            'expression_levels': expression_levels[:20],
                            'antibody_info': antibody_info[:10],
                            'source': 'HPA_XML'
                        }
                    else:
                        self.logger.warning(f"HPA XML returned no usable data for {gene_name}")
                        return None
            
                except ET.ParseError as e:
                    self.logger.debug(f"HPA XML parsing failed for {ensembl_id}: {e}")
                    return None
            else:
                self.logger.debug(f"HPA XML request failed for {ensembl_id}: HTTP {response.status_code}")
                return None
            
        except Exception as e:
            self.logger.debug(f"HPA XML query failed for {ensembl_id}: {e}")
            return None
        
    def _extract_location_name(self, element):
        """Extract location name from XML element"""
        try:
            # Try different ways to get the location name
        
            # Method 1: Check common attributes
            for attr in ['name', 'location', 'value', 'term']:
                if element.get(attr):
                    return element.get(attr).strip()
        
            # Method 2: Check text content
            if element.text and element.text.strip():
                return element.text.strip()
        
            # Method 3: Check first child element with text
            for child in element:
                if child.text and child.text.strip():
                    return child.text.strip()
        
            return None
        
        except Exception:
            return None
    
    def _get_hpa_web_safe(self, gene_name):
        """Safe HPA web scraping with corruption detection"""
        try:
            # Try HPA search API first
            search_url = "https://www.proteinatlas.org/api/search_download.php"
            params = {
                'search': gene_name,
                'format': 'json',
                'columns': 'g,gid,up,pe,ts'  # gene, gene_id, uniprot, protein_evidence, tissue_specificity
            }
            
            response = self.session.get(search_url, params=params, timeout=20)
            
            if response.status_code == 200:
                try:
                    data = response.json()
                    
                    tissues = set()
                    expression_levels = []
                    
                    for entry in data:
                        gene = entry.get('Gene', '')
                        if gene.upper() == gene_name.upper():
                            tissue_specificity = entry.get('Tissue specificity', '')
                            protein_evidence = entry.get('Protein evidence', '')
                            
                            # Parse tissue specificity
                            if tissue_specificity and tissue_specificity != 'n/a':
                                # Extract tissue names from tissue specificity description
                                tissue_matches = self._extract_tissues_from_text(tissue_specificity)
                                tissues.update(tissue_matches)
                                
                                for tissue in tissue_matches:
                                    expression_levels.append(f"{tissue}:detected")
                    
                    if tissues:
                        return {
                            'gene_name': gene_name,
                            'tissues': list(tissues)[:15],
                            'subcellular_locations': [],
                            'expression_levels': expression_levels[:15],
                            'antibody_info': [],
                            'source': 'HPA_API'
                        }
                
                except json.JSONDecodeError:
                    self.logger.debug(f"HPA API returned invalid JSON for {gene_name}")
            
            return None
            
        except Exception as e:
            self.logger.debug(f"HPA web query failed for {gene_name}: {e}")
            return None
    
    def _is_valid_location_name(self, location_name):
        """Validate if location name looks legitimate"""
        if not location_name or len(location_name) < 3:
            return False
    
        location_lower = location_name.lower().strip()
    
        # Check for obvious corruption patterns
        corruption_indicators = [
            'javascript', 'function', 'document', 'window', 'var ',
            'html', 'css', 'style', 'class=', 'id=', '{', '}',
            'position:', 'width:', 'height:', 'px', 'em',
            'humanproteome', 'proteinatlas'
        ]
    
        if any(indicator in location_lower for indicator in corruption_indicators):
            return False
    
        # Check if it's too long (likely corrupted)
        if len(location_name) > 50:
            return False
    
        # Check for reasonable cellular location terms
        valid_location_keywords = [
            'nucleus', 'cytoplasm', 'membrane', 'mitochondria', 'endoplasmic',
            'golgi', 'ribosome', 'lysosome', 'peroxisome', 'secreted',
            'extracellular', 'plasma', 'nuclear', 'cytoplasmic', 'vesicle',
            'organelle', 'compartment', 'reticulum', 'apparatus'
        ]
    
        # If it contains any valid keywords, it's probably good
        if any(keyword in location_lower for keyword in valid_location_keywords):
            return True
    
        # If it's short and alphanumeric, it might be valid
        if len(location_name) <= 30 and location_name.replace(' ', '').replace('-', '').isalnum():
            return True
    
        return False

    def _get_expression_atlas_enhanced(self, gene_name):
        """Enhanced Expression Atlas query"""
        try:
            url = "https://www.ebi.ac.uk/gxa/json/experiments"
            params = {
                'geneQuery': gene_name,
                'species': 'homo sapiens',
                'experimentType': 'baseline'
            }
            
            response = self.session.get(url, params=params, timeout=20)
            
            if response.status_code == 200:
                data = response.json()
                experiments = data.get('experiments', [])
                
                tissues = set()
                expression_levels = []
                
                for experiment in experiments[:15]:
                    exp_description = experiment.get('description', '').lower()
                    exp_type = experiment.get('experimentType', '')
                    
                    # Extract tissue information from experiment descriptions
                    tissue_matches = self._extract_tissues_from_text(exp_description)
                    
                    for tissue in tissue_matches:
                        tissues.add(tissue.title())
                        expression_levels.append(f"{tissue.title()}:expressed")
                
                if tissues:
                    return {
                        'gene_name': gene_name,
                        'tissues': list(tissues)[:15],
                        'subcellular_locations': [],
                        'expression_levels': expression_levels[:15],
                        'antibody_info': [],
                        'source': 'Expression_Atlas'
                    }
            
            return None
            
        except Exception as e:
            self.logger.debug(f"Expression Atlas enhanced query failed for {gene_name}: {e}")
            return None
    
    def _get_gtex_data(self, gene_name):
        """Get tissue expression data from GTEx (Genotype-Tissue Expression project)"""
        try:
            # GTEx API endpoint (if available) or alternative sources
            # Note: GTEx doesn't have a simple REST API, so this would need to be adapted
            # based on available GTEx data sources or downloads
            
            # For now, we'll use UniProt as a proxy for tissue expression
            return self._get_uniprot_tissue_enhanced(gene_name)
            
        except Exception as e:
            self.logger.debug(f"GTEx query failed for {gene_name}: {e}")
            return None
    
    def _get_uniprot_tissue_enhanced(self, gene_name):
        """Enhanced UniProt tissue specificity extraction"""
        try:
            url = "https://rest.uniprot.org/uniprotkb/search"
            params = {
                'query': f'gene:{gene_name} AND organism_id:9606',
                'format': 'json',
                'fields': 'accession,cc_function,cc_tissue_specificity,cc_developmental_stage',
                'size': 3
            }
            
            response = self.session.get(url, params=params, timeout=15)
            
            if response.status_code == 200:
                data = response.json()
                results = data.get('results', [])
                
                tissues = set()
                expression_levels = []
                
                for protein in results:
                    comments = protein.get('comments', [])
                    
                    for comment in comments:
                        comment_type = comment.get('commentType', '')
                        
                        if comment_type in ['TISSUE SPECIFICITY', 'FUNCTION', 'DEVELOPMENTAL STAGE']:
                            texts = comment.get('texts', [])
                            
                            for text in texts:
                                text_value = text.get('value', '')
                                
                                # Extract tissue mentions from text
                                tissue_matches = self._extract_tissues_from_text(text_value)
                                
                                for tissue in tissue_matches:
                                    tissues.add(tissue.title())
                                    expression_levels.append(f"{tissue.title()}:mentioned")
                
                if tissues:
                    return {
                        'gene_name': gene_name,
                        'tissues': list(tissues)[:15],
                        'subcellular_locations': [],
                        'expression_levels': expression_levels[:15],
                        'antibody_info': [],
                        'source': 'UniProt_Tissue'
                    }
            
            return None
            
        except Exception as e:
            self.logger.debug(f"UniProt tissue enhanced query failed for {gene_name}: {e}")
            return None
    
    def _extract_tissues_from_text(self, text):
        """Extract tissue names from descriptive text"""
        if not text:
            return []
        
        text_lower = text.lower()
        
        # Comprehensive tissue keyword mapping
        tissue_keywords = {
            'brain': 'Brain',
            'cerebral': 'Brain',
            'neuronal': 'Brain',
            'cortex': 'Brain cortex',
            'hippocampus': 'Hippocampus',
            'liver': 'Liver',
            'hepatic': 'Liver',
            'hepatocyte': 'Liver',
            'kidney': 'Kidney',
            'renal': 'Kidney',
            'heart': 'Heart',
            'cardiac': 'Heart',
            'myocardium': 'Heart muscle',
            'lung': 'Lung',
            'pulmonary': 'Lung',
            'muscle': 'Muscle',
            'skeletal muscle': 'Skeletal muscle',
            'smooth muscle': 'Smooth muscle',
            'skin': 'Skin',
            'dermal': 'Skin',
            'epidermis': 'Skin',
            'blood': 'Blood',
            'plasma': 'Blood',
            'serum': 'Blood',
            'bone': 'Bone',
            'skeletal': 'Bone',
            'testis': 'Testis',
            'ovary': 'Ovary',
            'ovarian': 'Ovary',
            'breast': 'Breast',
            'mammary': 'Breast',
            'prostate': 'Prostate',
            'pancreas': 'Pancreas',
            'pancreatic': 'Pancreas',
            'thyroid': 'Thyroid',
            'adrenal': 'Adrenal gland',
            'spleen': 'Spleen',
            'lymph': 'Lymphoid tissue',
            'intestine': 'Intestine',
            'colon': 'Colon',
            'stomach': 'Stomach',
            'gastric': 'Stomach',
            'esophagus': 'Esophagus',
            'trachea': 'Trachea',
            'bladder': 'Bladder',
            'uterus': 'Uterus',
            'placenta': 'Placenta',
            'eye': 'Eye',
            'retina': 'Retina',
            'cornea': 'Cornea'
        }
        
        found_tissues = set()
        
        for keyword, tissue_name in tissue_keywords.items():
            if keyword in text_lower:
                found_tissues.add(tissue_name)
        
        return list(found_tissues)
    
    def _is_valid_hpa_data(self, hpa_data):
        """Validate HPA data quality and detect corruption"""
        if not hpa_data:
            return False
        
        tissues = hpa_data.get('tissues', [])
        expression_levels = hpa_data.get('expression_levels', [])
        subcellular_locations = hpa_data.get('subcellular_locations', [])
        
        # Check if we have any meaningful data
        if not (tissues or expression_levels or subcellular_locations):
            return False
        
        # Check for corruption patterns
        corruption_indicators = [
            'humanproteome', 'position:', 'collision:', 'width:', 
            'javascript', 'function', '{', '}', 'var ', 'document.',
            'sub_section', 'appendelem', 'htmlelement', 'nodelist'
        ]
        
        # Check all data fields for corruption
        all_data = []
        all_data.extend(tissues)
        all_data.extend([str(item) for item in expression_levels])
        all_data.extend(subcellular_locations)
        
        for item in all_data:
            item_str = str(item).lower()
            for indicator in corruption_indicators:
                if indicator in item_str:
                    self.logger.debug(f"HPA data corruption detected: {indicator} in {item_str[:50]}")
                    return False
        
        return True
    
    def _is_valid_tissue_name(self, tissue_name):
        """Validate if tissue name looks legitimate"""
        if not tissue_name or len(tissue_name) < 3:
            return False
    
        tissue_lower = tissue_name.lower().strip()
    
        # Check for corruption
        corruption_indicators = [
            'javascript', 'function', 'document', 'humanproteome'
        ]
    
        if any(indicator in tissue_lower for indicator in corruption_indicators):
            return False
    
        # Check length
        if len(tissue_name) > 50:
            return False
    
        return True
    
    def _process_hpa_data_enhanced(self, results, idx, data, gene_name):
        """Process HPA data with correct column names and better subcellular location handling"""
        if not data or not self._is_valid_hpa_data(data):
            self._set_hpa_no_value(results, idx)
            return
    
        try:
            tissues = data.get('tissues', [])
            subcellular_locations = data.get('subcellular_locations', [])  # This should now have data
            expression_levels = data.get('expression_levels', [])
            antibody_info = data.get('antibody_info', [])
            source = data.get('source', 'Unknown')
        
            # Process primary tissue (highest priority/most mentioned)
            if tissues:
                # Count tissue mentions to find primary
                tissue_counts = {}
                for tissue in tissues:
                    tissue_counts[tissue] = tissue_counts.get(tissue, 0) + 1
            
                # Get primary tissue (most mentioned, or first if tie)
                primary_tissue = max(tissue_counts.keys(), key=lambda k: tissue_counts[k])
                results.at[idx, 'hpa_primary_tissue'] = primary_tissue
            
                # Determine expression level for primary tissue
                primary_expression = "detected"
                for expr in expression_levels:
                    if primary_tissue.lower() in expr.lower():
                        # Extract level from expression string
                        if ':' in expr:
                            level_part = expr.split(':', 1)[1]
                            primary_expression = level_part.strip()
                        break
            
                results.at[idx, 'hpa_expression_level'] = primary_expression
            else:
                results.at[idx, 'hpa_primary_tissue'] = "NO VALUE FOUND"
                results.at[idx, 'hpa_expression_level'] = "NO VALUE FOUND"
        
            # Process all tissue expression data
            if expression_levels:
                # Clean and format expression data for easy parsing
                clean_expressions = []
                for expr in expression_levels[:20]:  # Limit to top 20
                    clean_expr = str(expr).replace(f" ({source})", "")  # Remove source tags
                    clean_expressions.append(clean_expr)
                results.at[idx, 'hpa_all_tissues'] = " | ".join(clean_expressions)
            else:
                results.at[idx, 'hpa_all_tissues'] = "NO VALUE FOUND"
        
            # FIXED: Process subcellular locations with better handling
            if subcellular_locations:
                # Clean and validate locations
                clean_locations = []
                for loc in subcellular_locations[:10]:  # Limit to top 10
                    clean_loc = str(loc).strip()
                    if clean_loc and len(clean_loc) > 2:  # Basic validation
                        clean_locations.append(clean_loc)
            
                if clean_locations:
                    results.at[idx, 'hpa_subcellular_location'] = " | ".join(clean_locations)
                    self.logger.info(f"HPA subcellular locations for {gene_name}: {clean_locations}")
                else:
                    results.at[idx, 'hpa_subcellular_location'] = "NO VALUE FOUND"
                    self.logger.warning(f"HPA subcellular locations were found but failed validation for {gene_name}")
            else:
                results.at[idx, 'hpa_subcellular_location'] = "NO VALUE FOUND"
                self.logger.debug(f"No HPA subcellular locations found for {gene_name}")
        
            # Process antibody reliability as numerical score
            if antibody_info:
                # Convert reliability info to numerical score (0-5)
                reliability_score = self._calculate_reliability_score(antibody_info)
                results.at[idx, 'hpa_antibody_reliability'] = reliability_score
            else:
                results.at[idx, 'hpa_antibody_reliability'] = 3  # Default medium reliability
        
            # Store data source
            results.at[idx, 'hpa_data_source'] = source
        
            # Log success with subcellular location info
            subcell_info = "with subcellular locations" if subcellular_locations else "no subcellular locations"
            self.logger.info(f"HPA success for {gene_name}: Primary tissue={results.at[idx, 'hpa_primary_tissue']}, Total tissues={len(tissues)}, {subcell_info} ({source})")
        
        except Exception as e:
            self.logger.error(f"Error processing HPA data for {gene_name}: {e}")
            self._set_hpa_no_value(results, idx)
    
    def _calculate_reliability_score(self, antibody_info):
        """Convert antibody reliability information to numerical score (0-5)"""
        try:
            if not antibody_info:
                return 3  # Default medium
            
            # Convert reliability terms to scores
            reliability_text = " ".join(str(info).lower() for info in antibody_info)
            
            if any(term in reliability_text for term in ['high', 'excellent', 'validated']):
                return 5
            elif any(term in reliability_text for term in ['good', 'reliable']):
                return 4
            elif any(term in reliability_text for term in ['medium', 'moderate']):
                return 3
            elif any(term in reliability_text for term in ['low', 'poor']):
                return 2
            elif any(term in reliability_text for term in ['unknown', 'uncertain']):
                return 1
            else:
                return 3  # Default medium
        
        except Exception:
            return 3  # Default medium
    
    def _set_compartments_no_value(self, results, idx):
        """Set COMPARTMENTS fields to NO VALUE FOUND"""
        compartments_fields = [
            'compartments_primary_location', 'compartments_primary_confidence',
            'compartments_all_locations', 'compartments_confidence_scores',
            'compartments_evidence_types', 'compartments_data_sources'
        ]
        for field in compartments_fields:
            if field == 'compartments_primary_confidence':
                results.at[idx, field] = 0  # Numerical 0 for no confidence
            else:
                results.at[idx, field] = "NO VALUE FOUND"
    
    def _set_hpa_no_value(self, results, idx):
        """Set HPA fields to NO VALUE FOUND"""
        hpa_fields = [
            'hpa_primary_tissue', 'hpa_expression_level', 'hpa_all_tissues',
            'hpa_subcellular_location', 'hpa_antibody_reliability', 'hpa_data_source'
        ]
        for field in hpa_fields:
            if field == 'hpa_antibody_reliability':
                results.at[idx, field] = 0  # Numerical 0 for unknown reliability
            else:
                results.at[idx, field] = "NO VALUE FOUND"
    
    def _log_analysis_statistics(self, results, options):
        """Log success statistics for debugging"""
        total = len(results)
        
        if options.get('compartments', False):
            comp_success = sum(1 for _, row in results.iterrows() 
                             if row.get('compartments_location', 'NO VALUE FOUND') != 'NO VALUE FOUND')
            self.logger.info(f"Enhanced COMPARTMENTS final success rate: {comp_success}/{total} ({comp_success/total*100:.1f}%)")
        
        if options.get('hpa', False):
            hpa_success = sum(1 for _, row in results.iterrows() 
                            if row.get('hpa_tissue_expression', 'NO VALUE FOUND') != 'NO VALUE FOUND' and
                               'Humanproteome' not in str(row.get('hpa_tissue_expression', '')))
            self.logger.info(f"HPA final success rate: {hpa_success}/{total} ({hpa_success/total*100:.1f}%)")

    def _parse_location_from_text(self, text):
        """Parse subcellular locations mentioned in text"""
        if not text:
            return set()
    
        text_lower = text.lower()
        found_locations = set()
    
        # Common subcellular location terms
        location_terms = {
            'nucleus': 'Nucleus',
            'nuclear': 'Nucleus',
            'cytoplasm': 'Cytoplasm',
            'cytoplasmic': 'Cytoplasm',
            'mitochondria': 'Mitochondrion',
            'mitochondrial': 'Mitochondrion',
            'membrane': 'Membrane',
            'plasma membrane': 'Plasma membrane',
            'cell membrane': 'Cell membrane',
            'endoplasmic reticulum': 'Endoplasmic reticulum',
            'er': 'Endoplasmic reticulum',
            'golgi': 'Golgi apparatus',
            'ribosome': 'Ribosome',
            'ribosomal': 'Ribosome',
            'lysosome': 'Lysosome',
            'lysosomal': 'Lysosome',
            'peroxisome': 'Peroxisome',
            'secreted': 'Secreted',
            'extracellular': 'Extracellular'
        }
    
        for term, standard_name in location_terms.items():
            if term in text_lower:
                found_locations.add(standard_name)
    
        return found_locations

class EnsemblGeneMapper:
    """Enhanced Ensembl gene mapper with improved gene symbol handling"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.cache = {}
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'ProtMerge/1.2.0',
            'Accept': 'application/json'
        })
    
    def get_ensembl_id(self, gene_symbol):
        """Get Ensembl gene ID with enhanced gene symbol matching"""
        if gene_symbol in self.cache:
            return self.cache[gene_symbol]
        
        try:
            gene_symbol = str(gene_symbol).strip()
            if not gene_symbol:
                return None
            
            # Try multiple strategies
            strategies = [
                gene_symbol,           # Exact match
                gene_symbol.upper(),   # Uppercase
                gene_symbol.lower(),   # Lowercase
            ]
            
            # Add common gene symbol variations
            if gene_symbol.upper() != gene_symbol:
                strategies.append(gene_symbol.upper())
            
            # Try with and without common suffixes/prefixes
            if gene_symbol.endswith(('1', 'A', 'B')):
                strategies.append(gene_symbol[:-1])
            
            for strategy in strategies:
                ensembl_id = self._try_ensembl_direct(strategy)
                if ensembl_id:
                    self.cache[gene_symbol] = ensembl_id
                    return ensembl_id
            
            # Try alternative search using gene aliases
            ensembl_id = self._try_ensembl_search(gene_symbol)
            if ensembl_id:
                self.cache[gene_symbol] = ensembl_id
                return ensembl_id
            
            # Cache failures to avoid repeated lookups
            self.cache[gene_symbol] = None
            return None
            
        except Exception as e:
            self.logger.debug(f"Enhanced Ensembl lookup failed for {gene_symbol}: {e}")
            self.cache[gene_symbol] = None
            return None
    
    def _try_ensembl_direct(self, gene_symbol):
        """Direct Ensembl REST API lookup"""
        try:
            url = f"https://rest.ensembl.org/lookup/symbol/homo_sapiens/{gene_symbol}"
            
            response = self.session.get(url, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                ensembl_id = data.get('id')
                if ensembl_id and ensembl_id.startswith('ENSG'):
                    return ensembl_id
            
            return None
            
        except Exception:
            return None
    
    def _try_ensembl_search(self, gene_symbol):
        """Try Ensembl search API for gene aliases"""
        try:
            url = "https://rest.ensembl.org/lookup/symbol/homo_sapiens"
            params = {'expand': 1}
            
            # This is a simplified approach - in practice, you might need to use
            # the Ensembl search API or BioMart for alias searching
            return None
            
        except Exception:
            return None