# SharePointIFrame
A basic PCF for displaying the native SharePoint document library experience as an iframe in dataverse

## Code Overview
This control is implemented as a Virtual PCF component, using the React libraries from Dataverse. The control browses a given record's associated Document Locations and builds a dropdown with those location options. It climbs up the tree of locations until it reaches its first Site, then builds a URL and sets an IFrame src with that value. The dropdown and iframe are implmented with a Stack component, and may need to be rescaled if you have non-desktop browser applications in mind (the sizing has been hardcoded to keep implenentation simple, but can be parameterized).

## Contributing
Contributions are welcome! Just shoot me a pull request and I will be happy to review any suggestions you may have!

## Usage
Not interested in how the code works? Just want to skip the solution? go [here](https://github.com/cchannon/SharePointIFrame/tree/main/SolutionFile) to download the zip and install, then follow these installation instructions:

- Set up SharePoint Document Integration in your target environment and put a subgrid of the **SharePoint Document Locations** on your target form
- Import the Solution File to import the custom control into your environment
- From the form customizations area (at the time of writing this README you still need to use Classic interface to do this), select the properties of your subgrid and go to Controls (last tab)
- Select Add Control and add the **documentsIframe** control
- Enable the Control for web, mobile, and tablet
- Save and Publish your changes
