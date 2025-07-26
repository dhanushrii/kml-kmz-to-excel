# KML/KMZ to Excel Converter
This tool was specifically built for a real-world use case—to help my father convert KML/KMZ geospatial data into Excel format for his professional needs.  
It is **not a fully generalized tool**, but rather tailored for structured files that follow a specific format (e.g., consistent placemark structure, naming, etc.).

⚠️ **Note**: If you're using this for other KML/KMZ files, results may vary based on the structure of your input files.

**Features**
- Upload KML or KMZ files
- Extract placemark name, coordinates (latitude & longitude), and additional metadata
- Download output as an Excel spreadsheet
- Clean and user-friendly web interface

**Tech Stack**
- Frontend: HTML, CSS  
- Backend: Node.js, Express.js  
- python: xml, zipfile, openpyxl for file parsing and Excel generation  
