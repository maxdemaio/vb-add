# vb-add
**A Microsoft Excel Add-in to automate business processes**

## Overview

This is a repository for an Excel Add-in application called `vb-add`. Module1.bas contains all subroutines relating to data manipulation and filtering. Module2.bas contains a subroutine to display the application's information. The Ribbon file has vb-add as a tab, where its functions point to the subroutines within vb-add.xla. The vb-add.xla file is the final Add-in file that houses all the modules containing the subroutines.

## Installation

1. Clone this repository
2. Open `exampleFile.xlsx` and add vb-add to the ribbon by importing `vb-add-Ribbon.exportedUI`
    - File>Options>Customize Ribbon>Import/Export>Import customization file
3. Add the repository to 'Trusted Locations in the Trust Center Settings
    - File>Options>Trust Center>Trust Center Settings>Trusted Locations>Add new location