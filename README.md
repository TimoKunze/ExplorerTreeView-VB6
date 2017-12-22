# ExplorerTreeView-VB6
An ActiveX control for Visual Basic 6 that allows browsing for a folder in a tree view.

I've developed this ActiveX control between 2000 and 2002 and did update it on a regular basis until 2007. The last update has taken place in 2011. I don't intend to work on this project, but I think the code might be of some use to others.

# Before you make changes
If you make changes to the code and deploy the binary, keep in mind that ActiveX controls are COM components and therefore should stay binary compatible as long as you don't change the COM object's, i.e. the ActiveX control's public class name. Otherwise people using these components are likely to end up in the famous COM hell.
