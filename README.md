<p align=center>
  <a href="https://github.com/TimoKunze/ExplorerTreeView-VB6/releases/tag/1.15.3">
    <img alt="Release 1.15.3 Release" src="https://img.shields.io/badge/release-1.15.3-0688CB.svg">
  </a>
  <a href="https://github.com/TimoKunze/ExplorerTreeView-VB6/releases">
    <img alt="Download ExplorerTreeView" src="https://img.shields.io/badge/download-latest-0688CB.svg">
  </a>
  <a href="https://github.com/TimoKunze/ExplorerTreeView-VB6/blob/master/LICENSE">
    <img alt="License" src="https://img.shields.io/badge/license-MIT-0688CB.svg">
  </a>
  <a href="https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ExplorerTreeView&no_shipping=1&tax=0&currency_code=EUR">
    <img alt="Donate" src="https://img.shields.io/badge/%24-donate-E44E4A.svg">
  </a>
</p>

# ExplorerTreeView-VB6
An ActiveX control for Visual Basic 6 that allows browsing for a folder in a tree view.

I've developed this ActiveX control between 2000 and 2002 and did update it on a regular basis until 2007. The last update has taken place in 2011. I don't intend to work on this project, but I think the code might be of some use to others.

# Before you make changes
If you make changes to the code and deploy the binary, keep in mind that ActiveX controls are COM components and therefore should stay binary compatible as long as you don't change the COM object's, i.e. the ActiveX control's public class name. Otherwise people using these components are likely to end up in the famous COM hell.
