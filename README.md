# boardbuilder

# Installation
Refer to *install.txt* for the list of required packages.

# Usage
Source Excel file has to have **Title** and **Description** columns on each sheet.
Sheets not passing the check will be skipped by the builder.

Run *cardbuilder.py* with options:

**-s, --source:** Excel source file. **Mandatory**

**-o, --output:** Output folder. **Default:** execution directory.
 
**-f, --format:** pdf. **Default:** pdf.
 
**-t, --tabletop:** Export for Tabletop Simulator **TBD**

# Customization
It is possible to customize card Title and Description generation using **cust/cust_title.py** and **cust/cust_description.py files** respectively.
