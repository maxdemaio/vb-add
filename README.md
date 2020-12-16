# vb-add
**A Microsoft Excel Add-in to automate business processes**

![Project Preview](projPreview.png)

## Overview

This is a repository for an Excel Add-in application called vb-add. To checkout the code, `Module1.bas` contains all subroutines relating to data manipulation/filtering and `Module2.bas` contains a subroutine to display the application's information. The Ribbon file `vb-add-Ribbon.exportedUI` adds the application to your ribbon, where its functions point to the subroutines within `vb-add.xla`. The `vb-add.xla` file houses all the modules containing the subroutines.

## Installation

1. Clone this repository
2. Open `exampleFile.xlsx` and add vb-add to the ribbon by importing `vb-add-Ribbon.exportedUI`
    - File>Options>Customize Ribbon>Import/Export>Import customization file
3. Add the repository to Trusted Locations in the Trust Center Settings
    - File>Options>Trust Center>Trust Center Settings>Trusted Locations>Add new location