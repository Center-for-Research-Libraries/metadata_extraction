# Metadata Extraction
Extract metadata from Marc records via Folio and Worldcat.

## Basic Requirements

To install and use, you will need:

* Python (version 3.6 or later), with the ability to install extra Python libraries.
* Folio login for the Folio instance.
* Windows 10 or later

## Quick Start

1. Install Python 3. Add Python to your PATH.
2. Install the needed Python modules by typing `pip install -r requirements.txt`.
3. Click 'Settings' in the top menu then click 'Folio API'.
4. Fill out the 'Folio API' dialog box with your Folio username, password, OKAPI url, and tenant.
5. Click 'Apply' if login should only apply this session.  Click 'Save' if login should be perpetual.
6. 'Input' tab: choose appropriate data provider (Folio or Worldcat) and identifier via bottom left button.  Defaults to 'Folio (uuid)'.
7. 'Input' tab: place the identifier values (Folio uuids or OCLC number) in the first column, and place the ddsnext uuid values in the first column
8. Run via button or menu to retrieve marc records and extract metadata.
9. To export files, click 'File' then click 'Export'.
10. Select the file location for the descriptive metadata, title metadata, and Worldshare metadata.  Leave any unnecessary/unneeded blank.
11. Click 'Export' to create the output files.


