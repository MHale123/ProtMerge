# ProtMerge v1.2.0
**Protein Analysis Tool**

ProtMerge automatically gathers comprehensive protein information from multiple databases using UniProt IDs and outputs formatted Excel files.

## Quick Start

### Simple Installation (WINDOWS)
1. **Download this repository** (green "Code" button → "Download ZIP")
2. **Extract the folder** anywhere on your computer
3. 

The program will automatically install required packages (first time only) and launch ProtMerge.

### Requirements
- **Windows 10/11**
- **Python 3.8+** ([download from python.org](https://python.org))
- **Internet connection** (for protein database access)

### Manual Installation (Advanced Users)
If you prefer manual setup:
```bash
pip install -r requirements.txt
python protmerge.py
```

## How to Use
1. **Select File**: Browse and select your Excel file containing UniProt IDs
2. **Choose Column**: Click on the column containing UniProt IDs  
3. **Configure Options**: Select desired analyses (UniProt, ProtParam, BLAST)
4. **Start Analysis**: Click "START ANALYSIS" and monitor progress

The program will automatically fetch data from UniProt, ProtParam, and optionally BLAST, then save results to a professionally formatted Excel file with multiple sheets.

## Features
- **UniProt Data**: Organism, gene names, protein functions, sequences, environment detection
- **ProtParam Analysis**: Molecular weight, pI, GRAVY, extinction coefficients
- **Amino Acid Composition**: Detailed breakdown of all amino acids with atomic composition
- **BLAST Search**: Similar protein identification with statistical significance (optional)
- **Professional Excel Output**: Multiple formatted sheets, clickable hyperlinks, publication-ready results
- **Progress Tracking**: Real-time analysis progress with detailed status updates
- **Error Recovery**: Handling of network issues and data problems

## Important BLAST Usage Guidelines
**NCBI BLAST servers are a shared resource.** When using BLAST analysis, please follow these guidelines to ensure service availability for the entire community:

### Usage Limits
- **Daily limit**: Maximum 100 BLAST searches per 24 hours
- **Rate limiting**: Do not submit searches more than once every 10 seconds
- **Optimal timing**: Run large analyses (50+ proteins) on weekends or between 9 PM - 5 AM Eastern time

### Technical Requirements
- ProtMerge automatically complies with NCBI's 10-second delay between searches
- Each search is properly attributed with email and tool parameters
- Searches exceeding daily limits may be moved to slower queues or blocked

### Recommendations
- **Small projects**: Use BLAST option freely (under 50 proteins)
- **Large projects**: Consider running during off-peak hours or use NCBI's Stand-alone BLAST+ for 100+ proteins
- **Alternative**: For extensive BLAST analysis, consider NCBI's Docker Image or Elastic BLAST solutions

*Note: ProtMerge enforces these guidelines automatically, but users should be aware of these limitations when planning large-scale analyses.*

## Version History

### V1.2.0 (Current)
- **Major code refactoring**: Modular architecture for better maintainability
- **Enhanced analyzers**: Improved UniProt, ProtParam, and BLAST implementations
- **Professional Excel formatting**: Multiple sheets with beautiful styling and auto-sizing
- **Better error handling**: Comprehensive logging and graceful failure recovery
- **Environment detection**: Enhanced body location mapping for proteins
- **Progress improvements**: More detailed real-time status updates
- **New GUI**: New sleeker graphical user interface
- **Data Viewer**: Built in data viewer


### V1.1.2 (Previous)
- Fixed UI button sizing issues
- Added real-time progress bar with detailed status updates
- Improved exit handling and completion dialogs
- Enhanced user experience with threaded analysis

### V1.1.0
- Added modern graphical user interface
- Implemented file/column selection system
- Replaced fixed Excel template with flexible input options
- Added analysis configuration options

### V1.0.0
- Initial release
- Basic Excel sheet processing functionality
- Command-line interface only

## Troubleshooting

### "Python not found"
- Install Python from [python.org](https://python.org)
- **Important**: Check "Add Python to PATH" during installation

### "Package installation failed"
- Check internet connection
- Try running `START_HERE.bat` as administrator
- Ensure firewall isn't blocking pip

### Analysis errors
- Verify UniProt IDs are correct format (e.g., P04637)
- Check internet connection for database access
- Some proteins may not have complete data available

## Credits
**Created by:** Matthew Hale  
**UI Enhancements, Code refactoring:** Claude (Anthropic), ChatGPT (OpenAI), Gemini 2.5 (Google)

## License

This project is licensed under the [MIT License](LICENSE), **with the following restriction**:

- This software is intended for **academic and non-commercial use only**.  
- If you wish to use ProtMerge for commercial purposes, please contact the author for a separate license.

You are free to:
- Use, share, and modify the code
- As long as it's **not for commercial profit or sale**

We ask that you:
- Cite this work in academic publications
- Respect the academic intent of this software

