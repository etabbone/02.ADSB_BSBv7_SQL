// --------------------------------------------------------
//
// This file is to configure the configurable settings.
// Load this file before script.js file at gmap.html.
//
// --------------------------------------------------------

// -- Output Settings -------------------------------------
// Show metric values
Metric = true; // true or false

// -- Map settings ----------------------------------------
// The Latitude and Longitude in decimal format
CONST_CENTERLAT = -15.869118;
CONST_CENTERLON = -47.920883;
// The google maps zoom level, 0 - 16, lower is further out
CONST_ZOOMLVL   = 13;

// -- Marker settings -------------------------------------
// The default marker color
MarkerColor	  = "rgb(225, 225, 0)";
SelectedColor = "rgb(204, 255, 255)";

// -- Site Settings ---------------------------------------
SiteShow    = true; // true or false
// The Latitude and Longitude in decimal format
SiteLat     = -15.869118;
SiteLon     = -47.920883;

SiteCircles = true; // true or false (Only shown if SiteShow is true)
// In nautical miles or km (depending settings value 'Metric')
SiteCirclesDistances = new Array(0.250,0.500,0.750,1,1.5,2,2.5,5,10,15);

