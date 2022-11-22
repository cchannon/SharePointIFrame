# SharePointIFrame
A basic PCF for displaying the native SharePoint document library experience as an iframe in dataverse

## Usage
Not interested in how the code works? Just want to skip the solution? go here to download the zip and install, then follow the installation instructions here.

## Code Overview
This control is implemented as a Virtual PCF component, using the React libraries from Dataverse. The control browses a given record's associated Document Locations and builds a dropdown with those location options. It climbs up the tree of locations until it reaches its first Site, then builds a URL and sets an IFrame src with that value. The dropdown and iframe are implmented with a Stack component, and may need to be rescaled if you have non-desktop browser applications in mind (the sizing has been hardcoded to keep implenentation simple, but can be parameterized).

## Contributing
Contributions are welcome! Just shoot me a pull request and I will be happy to review any suggestions you may have!