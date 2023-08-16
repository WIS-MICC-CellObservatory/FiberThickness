#@ String(choices=("singleFile", "wholeFolder", "AllSubFolders"), style="list", persist=true) processMode
#@ String (label="File Extension", persist=true, description="eg .tif") fileExtention
#@ String (choices=("Image width", "Pixel width"), persist=true, discription="Choose how the image will be scaled") SetScale
#@ String (label="Scale Value", persist=true, discription="Choose scaling value") ScaleVal
#@ Integer (label="Max fiber width (um)", persist=true, description ="For display and background subtraction") MaxEstimatedWidthUm
#@ String (label="Auto Threshold Method", choices={"Li", "Otsu", "Huang", "Moments", "IJ_IsoData", "IsoData"}, style="listBox", Value="Li", persist=true, description="Pixel Classifier Resolution, use Other and Custom Class Resolotion below for custom resolution") AutoThresholdMethod
#@ Integer (label="Num of Dilate/Erode operations before Fill Holes", Value=40, persist=true, description ="This help closing broken fibers, but may add limit the possibility to discard thin attached fibers") nDilateErode
#@ Boolean (label="limit Size of Holes to be closed", Value=true, persist=true, description ="To avoid attachment of small fibrils segments") limitCloseHoleSize
#@ Integer (label="Max Size of Holes to close (um2)", Value=100, persist=true, description ="To avoid attachment of small fibrils segments") MaxHoleSizeToClose
//#@ Integer (label="Min Fiber segment Area (um2)", Value=2000, persist=true, description ="To avoid attachment of small fibrils segments") MinFilmSegmentArea
//#@ Float (label="Min Size of Skeleton Segment (um2)", Value=25, persist=true, description ="to discard small fiber segments, note that it can remove segments from fragmented fibers, \nand that it is area of thin center line and not the thick fiber segment ") MinSkeletonArea
#@ Boolean (label="Keep Only Largest Skeleton Segment", Value=true, persist=true, description ="help to discard parallel segments that are due to artifacts") keepOnlyLargestSkeletonSegment
#@ Boolean (label="Stop at every step for debugging", Value=false, persist=false, description ="to see the effect ofchanging parameters at each individual step") debugFlag
#@ Boolean batchModeFlag

/*
 * FilmThickness.ijm
 * 
 * Measure Film thichness by measuring Local Thickness along the center line of the film
 * The macro allow procewssing of single file or whole folder. 
 * 
 * Workflow
 * ========
 * - Scale the image to correct units (assuming the width of the image is equal to imRealLength um)
 * - Segment the film into binary image
 * 		- Apply background subtraction using rolling ball - this is optional step , that is used by default. 
 * 		  for this it is important to set MaxEstimatedWidthUm correctly. the rolling ball size is set to MaxEstimatedWidthUm+10 
 * 		  it is especially important when you use AutoThreshold. 
 * 		- Use either AutoThreshold (Li) or fixed threshold (0-MaxFilmIntensity) 
 * 		  note that the fixed threshold values need to change if you use background subtraction
 * - Measure the local thickness of the binary image
 * - Find center line of the film 
 * 	> Erode the binary film - to achieve better skeletonization, 
 * 	  nErode value can be set in 3 ways controled by setErodeValStrategy: 
 * 	  - fixed value (nDefaultErode)
 * 	  - the maximum between (Mean LocThick - 2* Std LocThk)/3  && MinLocThick/3
 * 	  - fixed value or the above value if the default fixed value is smaller
 * 	  by default use nDefaultErode (=20 pixels), but lower this value to (Mean LocThick - 2* Std LocThk)/3 for thinner films 
 * 	> Skeletonize
 * 	> Discard skeleton areas that are up to ignoreBorderSize percent from the image border, 
 * 	  if ignoreTopDown=0 , only side borders are ignored, =1: both side and top/bottom borders of the image are ignored
 * 	> Discard (most of) the perpendicular skeleton segments by: 
 * 		= breaking the skeleton at branch points
 * 		= discarding skeleton segments smaller than MinSkeletonArea (usually perpendicular)
 * 		= optionally discard segments that are oriented between minAngleToFilter and maxAngleToFilter 
 * 		  NOTE: it is valid to use the above only if the films were always imaged at horizontal orientation 
 * 		  this is controled by filterByAngle
 * 	- Measure Mean/Median/Modal/Min/Max of Local thickness along the remaining skeleton segments,  
 * 	  which are assumed to be good approximation for the centerline 
 * - write summary line 
 * - save quality control images with overlay of the skeleton segments used for measurement, 
 *   on top of the original image and the local thickness image
 * 
 * - Add Mean/Std/Min/Max lines for the summary table
 * 
 * Usage
 * =====
 * 
 * NOTE: Before running the macro, go through the images and discard images that do not really contain film segments, 
 *       so that they will not be included in the folder statistics
 * 
 * - Drag and drop the macro to Fiji, click Run
 * - Set Process mode to singleFile or wholeFolder
 * - select a file to process, if wholeFolder is selected: all the files in the folder of the selected file are processed
 * - If batchModeFlag is selected (recomended) the macro will run faster and not dispalyed temporary images
 * 
 * - NOTE: It is very important to inspect All quality control images to verify that segmentation and centerline are correct 
 * 
 * Output
 * ======
 * For each input image FN, the following output files are saved in ResultsSubFolder under the input folder
 * - FN_OrigOverlay.tif 	- the original image with overlay of the segmented film in blue
 * - FN_LocThk.tif			- local thickness measurement (values are in um), you can inspect them if you open the file in Fiji
 * 	  						  Maximum displayed value is set to MaxEstimatedWidthUm for all images to alloow visual comparison
 * - FN_LocThkOverlay.tif 	- local thickness with overlay of the skeleton segments
 * - FN_SkelSegmentsResults.xls - the detailed measurements for each skeleton segment in the image  
 * - FN_SkelSegmentsRoiSet.zip  - the skeleton segments used for measurements
 * 
 *  Overlay colors can be controled by BoundaryColor and CenterlineColor
 * 
 * Summary.xls  - Table with one line for each input image files with average values of Mean and Median LocTchikness
 * FilmThicknessParameters.txt - Parameters used during the latest run
 * 
 * Dependencies
 * ============
 * Fiji with ImageJ version > 1.52s (Check Help=>About ImageJ, and if needed use Help=>Update ImageJ...
 * It is based on LocalThickness ImageJ plugin: https://imagej.net/Local_Thickness, by Bob Dougherty
 * Please cite Fiji (https://imagej.net/Citing) and LocalThickness if you use it for publication
 * 
 * By Ofra Golani, MICC Cell Observatory, Weizmann Institute of Science, March 2020
 * 
 * v6:
 * - move MaxEstimatedWidthUm to GUI
 * - move AutoThresholdMethod to GUI
 * - move nDilateErode to GUI 
 * - add maxHolesToClose and add to GUI
 * - add limitCloseHoleSize and add to GUI
 * - add keep only largest skeleton object - to help discard disconnected skel objects on different vertical position - this can be used only if the fibers are horizontaly oriented 
 * - add parameter ValidErodeFactor instead of fixed factor of 3 (1/3) that used before and changed it 4 (1/4) - to avoid breaking the skeletonized fiber 
 */
 
// ============ Parameters =======================================
var macroVersion = "v6";
//var fileExtention = ".tif";
//var imRealLength = 312 //550; //625 for 20x 1x (0.20 um/pixel); 312 for 40x 1x (0.11 um/pixel) //um
var SubstractBackgroundVal = MaxEstimatedWidthUm + 10; //100; //um - should be larger than the maximum expected fiber thickness, 
								// if <=0 don't apply,  desired especially for autoThreshold
var SegmentMethod = "AutoThreshold"; // "FixThreshold"; // "AutoThreshold"
var MaxFilmIntensity = 8000; //6000; // Fixed Threshold value for film segmentation 
//var AutoThresholdMethod = "Li";
//var	nDilateErode = 40; // pixels - to assist with Fill Holes close to border
//var limitCloseHoleSize = 1;
//var MaxHoleSizeToClose = 100; // um^2
var MinFilmSegmentArea = 2000; //5000; // um^2

var setErodeValStrategy = "AlwaysUpdate"; // "Fix", "FixAndUpdate", "AlwaysUpdate"
var	ValidErodeFactor = 3; //4; // 4 is actually 1/4 , was 3 which is actually 1/3
var nDefaultErode = 20; //30; //50; // pixels - Try value that is about 1/3 of film width 
var MinSkeletonArea = 0.1; //5; //um^2
var MaxLocThkDisplayValue = MaxEstimatedWidthUm; //90; //55;

var filterByAngle = 1;
var minAngleToFilter = 30; //40;
var maxAngleToFilter = 150; //140;

//var ignoreBorderSize = 30; //um , if value <=0 don't ignore border values 
var ignoreBorderSize = 5; // percentage of imRealLength eg 5 = 5%, if value <=0 don't ignore border values 
var ignoreTopDown = 0; 	  // 0=ignore only side borders, 1=ignore both side and top/bottom borders of the image
var ignoreFailedImagesFlag = 0;

var BoundaryColor = "blue";
var CenterlineColor = "green";

var ResultsSubFolder = "Results";
var cleanupFlag = 1; 
//var debugFlag = 0;

// Global Parameters
var SummaryTable = "SummaryResults.xls";
var nErode = nDefaultErode; 
var TimeString;

// ================= Main Code ====================================

Initialization();

// Choose image file or folder
if (matches(processMode, "singleFile")) {
	file_name=File.openDialog("Please select an image file to analyze");
	directory = File.getParent(file_name);
	}
else if (matches(processMode, "wholeFolder")) {
	directory = getDirectory("Please select a folder of images to analyze"); }

else if (matches(processMode, "AllSubFolders")) {
	parentDirectory = getDirectory("Please select a Parent Folder of subfolders to analyze"); }

if (batchModeFlag)
{
	print("Working in Batch Mode, processing without opening images");
	setBatchMode(true);
}

// Analysis 
if (matches(processMode, "wholeFolder") || matches(processMode, "singleFile")) {
	resFolder = directory + File.separator + ResultsSubFolder + File.separator; 
	File.makeDirectory(resFolder);
	print("inDir=",directory," outDir=",resFolder);
	SavePrms(resFolder);
	
	if (matches(processMode, "singleFile")) {
		ProcessFile(directory, resFolder, file_name); }
	else if (matches(processMode, "wholeFolder")) {
		ProcessFiles(directory, resFolder); }
}

else if (matches(processMode, "AllSubFolders")) {
	list = getFileList(parentDirectory);
	for (i = 0; i < list.length; i++) {
		if(File.isDirectory(parentDirectory + list[i])) {
			subFolderName = list[i];
			//print(subFolderName);
			subFolderName = substring(subFolderName, 0,lengthOf(subFolderName)-1);
			//print(subFolderName);

			//directory = parentDirectory + list[i];
			directory = parentDirectory + subFolderName + File.separator;
			resFolder = directory + ResultsSubFolder + File.separator; 
			//print(parentDirectory, directory, resFolder);
			File.makeDirectory(resFolder);
			print("inDir=",directory," outDir=",resFolder);
			SavePrms(resFolder);
			//if (isOpen("SummaryResults.xls"))
			if (isOpen(SummaryTable))
			{
				//selectWindow("SummaryResults.xls");
				selectWindow(SummaryTable);
				run("Close");  // To close non-image window
			}
			ProcessFiles(directory, resFolder);
			print("Processing ",subFolderName, " Done");
		}
	}
}

setBatchMode(false);
print("Done !");

// ================= Helper Functions ====================================

//--------------------------------------
// Loop on all files in the folder and Run analysis on each of them
function ProcessFiles(directory, resFolder) 
{
	dir1=substring(directory, 0,lengthOf(directory)-1);
	idx=lastIndexOf(dir1,File.separator);
	subdir=substring(dir1, idx+1,lengthOf(dir1));

	// Get the files in the folder 
	fileListArray = getFileList(directory);
	
	// Loop over files
	for (fileIndex = 0; fileIndex < lengthOf(fileListArray); fileIndex++) {
		if (endsWith(fileListArray[fileIndex], fileExtention) ) {
			file_name = directory+File.separator+fileListArray[fileIndex];
			//open(file_name);	
			//print("\nProcessing:",fileListArray[fileIndex]);
			showProgress(fileIndex/lengthOf(fileListArray));
			ProcessFile(directory, resFolder, file_name);
		} // end of if 
	} // end of for loop

	// Save Results
	if (isOpen(SummaryTable))
	{
		GenerateSummaryLines(SummaryTable);
		selectWindow(SummaryTable);
		SummaryTable1 = replace(SummaryTable, ".xls", "");
		print("SummaryTable=",SummaryTable,"SummaryTable1=",SummaryTable1,"subdir=",subdir);
		saveAs("Results", resFolder+SummaryTable1+"_"+subdir+".xls");
		run("Close");  // To close non-image window
	}
	
	// Cleanup
	if (cleanupFlag==true) {
		if (isOpen(SummaryTable))
		{
			selectWindow(SummaryTable);
			run("Close");  // To close non-image window
		}
	}
} // end of ProcessFiles



//--------------------------------------
// Run analysis of single file
function ProcessFile(directory, resFolder, file_name) 
{
	nErode = nDefaultErode;
	
	// ===== Open File ========================
	print("Processing:",file_name, "...");
	open(file_name);

	origName = getTitle();
	origNameNoExt = replace(origName, fileExtention, "");
	origNameNoExt = replace(origNameNoExt, " ", "_");
	origNameNoExt = replace(origNameNoExt, " ", "_");
	origNameNoExt = replace(origNameNoExt, " ", "_");
	
	// Set Scale
	//imWidthPixels = getWidth();
	//run("Set Scale...", "distance="+imWidthPixels+" known="+imRealLength+" unit=um");

	// Set Scale - imRealLength or LenPixRatio
	if (matches(SetScale, "Image width")) {
		imWidthPixels = getWidth(); 
		run("Set Scale...", "distance="+imWidthPixels+" known="+ScaleVal+" unit=um");
	}
	else if (matches(SetScale, "Pixel width")) {
		//run("Set Scale...", "distance=1 known="+ScaleVal" unit=um");
		//run("Properties...", "channels=1 slices=1 frames=1 pixel_width="+ScaleVal+" pixel_height="+ScaleVal+" voxel_depth=1");
		setVoxelSize(ScaleVal, ScaleVal, 1, "um");
	}

	// Segment the Film / Fiber
	//==================================
	selectWindow(origName);
	depth = bitDepth();
	if (depth == 24) // RGB Image 
		run("8-bit");

	run("Duplicate...", "title=Binary");
	if (SubstractBackgroundVal > 0)
	{
		getPixelSize(unit, pixelWidth, pixelHeight);
		rollingBallSize = SubstractBackgroundVal / pixelWidth; // in pixels
		print(origName,", rollingBallSize=",rollingBallSize);
		if(debugFlag) waitForUser("Before Subtract Background");		
		run("Subtract Background...", "rolling="+rollingBallSize+" light");
		if(debugFlag) waitForUser("After Subtract Background");		
	}
	if (matches(SegmentMethod, "FixThreshold"))
	{
		setThreshold(0, MaxFilmIntensity);
	}
	else if (matches(SegmentMethod, "AutoThreshold"))
	{
		setAutoThreshold(AutoThresholdMethod);
	}
	run("Convert to Mask");	
	if(debugFlag) waitForUser("Before DilateErode");
	for (m=0;m<nDilateErode;m++) run("Dilate");
	for (m=0;m<nDilateErode;m++) run("Erode");
	if(debugFlag) waitForUser("Before Fill Holes");

	if (limitCloseHoleSize)
	{
		// Close small holes
		run("Duplicate...", "title=FiberMasked_Tmp");
		run("Invert");
		run("Analyze Particles...", "size=0-"+MaxHoleSizeToClose+" show=Masks");
		selectWindow("Mask of FiberMasked_Tmp");
		rename("FiberMasked_SmallHoles");
		run("Invert LUT");
		imageCalculator("Add", "Binary","FiberMasked_SmallHoles");		
	} else 
		run("Fill Holes");

	selectWindow("Binary");
	run("Analyze Particles...", "size="+MinFilmSegmentArea+"-Infinity show=Masks");
	run("Invert LUT");
	rename("BinaryMask");
	
	run("Duplicate...", "title="+origNameNoExt+"_Thick");
	skelName = origNameNoExt+"_Skel";
	run("Duplicate...", "title="+skelName);
	
	// Calculate Local Thickness
	//============================
	selectWindow(origNameNoExt+"_Thick");
	run("Local Thickness (masked, calibrated, silent)");
	LocThkName = origNameNoExt+"_Thick_LocThk";

	//selectWindow("Binary");
	selectWindow("BinaryMask");
	run("Create Selection");
	selectWindow(LocThkName);
	run("Restore Selection");	
	if (matches(setErodeValStrategy, "AlwaysUpdate") ||  matches(setErodeValStrategy, "FixAndUpdate") )
	{
		getRawStatistics(nPixels, meanVal, minVal, maxVal, stdVal, histogram);
		getVoxelSize(width, height, depth, unit);
		maxValidErode1 = floor((meanVal/width-2*stdVal/width)/ValidErodeFactor) + 1 ; 
		maxValidErode2 = floor((minVal/width)/ValidErodeFactor) + 1 ; 
		maxValidErode = maxOf(maxValidErode1, maxValidErode2);
		maxValidErode = maxOf(maxValidErode, 0);
		if (matches(setErodeValStrategy, "AlwaysUpdate") )
		{ // always update
				nErode = maxValidErode; 
		}
		if (matches(setErodeValStrategy, "FixAndUpdate") )
		{ // update only if fixed value is too big
			if (nErode > maxValidErode)
				nErode = maxValidErode; 
		}
		if (debugFlag == 1)
			print(file_name, meanVal, minVal, maxVal, stdVal, meanVal/width, minVal/width, maxVal/width, stdVal/width, maxValidErode2, maxValidErode1, nErode);
	}
	print(origNameNoExt, ", setErodeValStrategy=", setErodeValStrategy,", nErode=", nErode);
	
	// Get (possibly broken) centerline 
	//==================================
	selectWindow(skelName);
	if(debugFlag) waitForUser("Before Erosion prior to skeletonization");
	// Erode to allow for better skelotinazation
	for (n=0; n<nErode; n++) run("Erode");
	if(debugFlag) waitForUser("After Erosion prior to skeletonization");
	run("Skeletonize");

	if(debugFlag) waitForUser("Before Keep only largest skeleton object");
	// Keep only larger Skeleton segment - to remove parallel skeleton segments that are due to artifacts
	if (keepOnlyLargestSkeletonSegment)
	{
		run("Connected Components Labeling", "connectivity=8 type=[8 bits]");
		run("Keep Largest Label");
		setOption("BlackBackground", true);
		run("Convert to Mask");
		skelName = getTitle();
	}
	DiscardBorderSkeleton(skelName, ignoreBorderSize, ignoreTopDown);
	 
	if(debugFlag) waitForUser("Before discarding by length and angle");
	// Disconnect perpendicular branches, by finding connection points in the skeleton and deleting them
	// use Branch Info table to get branch point location and set them to zero
	selectWindow(skelName);
	getVoxelSize(width, height, depth, unit);
	run("Set Scale...", "distance=0 known=0 unit=pixel");
	run("Analyze Skeleton (2D/3D)", "prune=[shortest branch] show");
	CloseTable("Results");
	IJ.renameResults("Branch information", "Results");
	for (n = 0; n < nResults ; n++)
	{
		v1x = getResult("V1 x", n);
		v1y = getResult("V1 y", n);
		v2x = getResult("V2 x", n);
		v2y = getResult("V2 y", n);
	
		ZeroPixelAndNeighbors(v1x, v1y);
		ZeroPixelAndNeighbors(v2x, v2y);
	}
	// set scale back
	setVoxelSize(width, height, depth, unit);
	
	// Remove short branches from analysis and measure the thickness from the LocThkGray image
	run("Set Measurements...", "area mean centroid standard modal min fit median display add redirect="+LocThkName+" decimal=2");
	//run("Analyze Particles...", "size="+MinSkeletonArea+"-Infinity show=Masks display clear summarize add");
	if ((ignoreFailedImagesFlag == 1) && (isOpen("Tagged Skeleton"))) // indication that something went wrong 
	{
		if (isOpen("Results"))
		run("Close");

		// Output the measured values into new results table
		if (isOpen(SummaryTable))
		{
			IJ.renameResults(SummaryTable, "Results"); // rename to avoid table overwrite
		}	
		else
			run("Clear Results");
	
		setResult("Label", nResults, origNameNoExt); 
		setResult("Area", nResults-1, 0); 
		
		// Save Results - actual saving is done at the higher level function as this table include one line for each image
		IJ.renameResults("Results", SummaryTable); // rename to avoid table overwrite
		if(cleanupFlag) Cleanup();
		return;
	}
	selectWindow(skelName);
	run("Analyze Particles...", "size="+MinSkeletonArea+"-Infinity show=Masks display clear add");
	rename("CleanSkeleton");

	selectWindow("Results");
	AngleArr = Table.getColumn("Angle");
	nSkeletonSegments = roiManager("count");
	if ( (nSkeletonSegments > 0) && (filterByAngle) )
	{
		for (n = nSkeletonSegments-1; n >= 0 ; n--)
		{
			//Angle = getResult("Angle", n);
			//if ( (Angle > minAngleToFilter) && (Angle < maxAngleToFilter) )
			if ( (AngleArr[n] > minAngleToFilter) && (AngleArr[n] < maxAngleToFilter) )
			{
				roiManager("select", n);
				roiManager("delete");
			}
		}
		run("Clear Results");
		roiManager("deselect");
		roiManager("measure");
	}

	// save Results 
	if (roiManager("count") > 0)
	{
		selectWindow("Results");
		Table.save(resFolder+origNameNoExt+"_SkelSegmentsResults.xls");
		roiManager("Deselect");
		roiManager("Save", resFolder+origNameNoExt+"_SkelSegmentsRoiSet.zip");
	// =========== Add line in Summary Table =============
		if (roiManager("count") > 1)
		{
			roiManager("Deselect");
			roiManager("Combine");
		}
		else 
			roiManager("select", 0);
			
		run("Measure");	
		Area = getResult("Area", nResults-1);
		Mean = getResult("Mean", nResults-1);
		Mode = getResult("Mode", nResults-1);
		Median = getResult("Median", nResults-1);
		Min = getResult("Min", nResults-1);
		Max = getResult("Max", nResults-1);
	}
	else // empty list 
	{
		Area = 0;
		Mean = 0;
		Mode = 0;
		Median = 0;
		Min = 0;
		Max = 0;	
	}
	run("Set Measurements...", "area mean standard modal min fit median display add redirect=None decimal=2");
	
	if (isOpen("Results"))
		run("Close");

	// Output the measured values into new results table
	if (isOpen(SummaryTable))
	{
		IJ.renameResults(SummaryTable, "Results"); // rename to avoid table overwrite
	}	
	else
		run("Clear Results");

	setResult("Label", nResults, origNameNoExt); 
	setResult("Area", nResults-1, Area); 
	setResult("MeanLocThk", nResults-1, Mean); 
	setResult("ModeLocThk", nResults-1, Mode); 
	setResult("MedianLocThk", nResults-1, Median); 
	setResult("MinLocThk", nResults-1, Min); 
	setResult("MaxLocThk", nResults-1, Max); 
	
	// Save Results - actual saving is done at the higher level function as this table include one line for each image
	IJ.renameResults("Results", SummaryTable); // rename to avoid table overwrite

	// Create Overlay images for Quality control 
	selectWindow(origName);
	roiManager("Set Color", CenterlineColor);
	roiManager("Set Line Width", 2);	
	roiManager("Show All without labels");
	run("Flatten");
	im=getImageID();
	
	// Save Local Thickness overlay image 
	selectWindow(LocThkName);
	setMinAndMax(0, MaxLocThkDisplayValue);
	run("Calibration Bar...", "location=[Upper Right] fill=White label=Black number=5 decimal=0 font=12 zoom=3 overlay");	
	roiManager("Show All without labels");
	saveAs("Tiff", resFolder+origNameNoExt+"_LocThk");
	run("Calibration Bar...", "location=[Upper Right] fill=White label=Black number=5 decimal=0 font=12 zoom=3");	
	run("Flatten");
	saveAs("Tiff", resFolder+origNameNoExt+"_LocThkOverlay");
	
	// Save orig overlay image 
	//selectWindow("Binary");
	selectWindow("BinaryMask");
	run("Create Selection");
	run("Properties... ", "  stroke="+BoundaryColor+" width=2");
	selectImage(im);
	run("Restore Selection");
	run("Flatten");
	saveAs("Tiff", resFolder+origNameNoExt+"_OrigOverlay");
		
	if(cleanupFlag) Cleanup();
} // end of ProcessFile


//===============================================================================================================
// mask out parts of skeleton that are up to ignoreBorderSize (um) from the image border
function DiscardBorderSkeleton(skelName, ignoreBorderSize, ignoreTopDown)
{
	if (ignoreBorderSize >= 0)
	{
		selectWindow(skelName);

		// Exclusion rectangle 
		getPixelSize(unit, pixelWidth, pixelHeight);
		//eDist=(ignoreBorderSize*imRealLength/100)/pixelWidth;
		//print("pixelWidth=",pixelWidth," eDist=",eDist);
		eDist=(ignoreBorderSize*getWidth/100); 
		print("Image width=",getWidth+" pixels"," eDist=",eDist);		
		if (ignoreTopDown == 1)
			makeRectangle(eDist, eDist, getWidth-2*eDist, getHeight-2*eDist);
		else 
			makeRectangle(eDist, 0, getWidth-2*eDist, getHeight);
		run("Create Mask");
		selectWindow(skelName);
		run("Select None");
		imageCalculator("AND", skelName,"Mask");
		//imageCalculator("AND create", skelName,"Mask");
	}
}

//===============================================================================================================
function ZeroPixelAndNeighbors(x, y)
{
	setPixel(x, y, 0);
	setPixel(x-1, y, 0);
	setPixel(x+1, y, 0);
	setPixel(x, y-1, 0);
	setPixel(x-1, y-1, 0);
	setPixel(x+1, y-1, 0);
	setPixel(x, y+1, 0);
	setPixel(x-1, y+1, 0);
	setPixel(x+1, y+1, 0);
	neighb5 = 1;
	if (neighb5 == 1)	
	{
		setPixel(x-2, y, 0);
		setPixel(x+2, y, 0);
		setPixel(x, y-2, 0);
		setPixel(x-2, y-2, 0);
		setPixel(x+2, y-2, 0);
		setPixel(x, y+2, 0);
		setPixel(x-2, y+2, 0);
		setPixel(x+2, y+2, 0);		

		setPixel(x-2, y-1, 0);
		setPixel(x+2, y-1, 0);
		setPixel(x-2, y+1, 0);
		setPixel(x+2, y+1, 0);

		setPixel(x-2, y-2, 0);
		setPixel(x+2, y-2, 0);
		setPixel(x-2, y+2, 0);
		setPixel(x+2, y+2, 0);
	}
}

//===============================================================================================================
function Initialization()
{
	run("Close All");
	run("Options...", "iterations=1 count=1 black");
	run("Set Measurements...", "area redirect=None decimal=3");
	CloseTable("Results");
	CloseTable(SummaryTable);
	roiManager("Reset");
	print("\\Clear");
}

//===============================================================================================================
function Cleanup()
{
	run("Close All");
	run("Clear Results");
	roiManager("reset");
	run("Collect Garbage");
	CloseTable("Branch information");
}


//===============================================================================================================
function CloseTable(TableName)
{
	if (isOpen(TableName))
	{
		selectWindow(TableName);
		run("Close");
	}
}

//===============================================================================================================
function SavePrms(resFolder)
{
	// print parameters to Prm file for documentation
	PrmFile = resFolder+"FilmThicknessParameters.txt";
	File.saveString("macroVersion="+macroVersion, PrmFile);
	File.append("", PrmFile); 
	setTimeString();
	File.append("RunTime="+TimeString, PrmFile)
	File.append("processMode="+processMode, PrmFile); 
	File.append("fileExtention="+fileExtention, PrmFile); 
	File.append("SetScale="+SetScale+" \n", PrmFile); 
	File.append("ScaleVal="+ScaleVal+" \n", PrmFile);
	File.append("MaxEstimatedWidthUm="+MaxEstimatedWidthUm, PrmFile); 
	File.append("SubstractBackgroundVal="+SubstractBackgroundVal, PrmFile); 
	File.append("SegmentMethod="+SegmentMethod, PrmFile); 
	File.append("MaxFilmIntensity="+MaxFilmIntensity, PrmFile); 
	File.append("AutoThresholdMethod="+AutoThresholdMethod, PrmFile); 
	File.append("nDilateErode="+nDilateErode, PrmFile); 
	File.append("limitCloseHoleSize="+limitCloseHoleSize, PrmFile); 
	File.append("MaxHoleSizeToClose="+MaxHoleSizeToClose, PrmFile); 
	File.append("MinFilmSegmentArea="+MinFilmSegmentArea, PrmFile); 
	File.append("setErodeValStrategy="+setErodeValStrategy, PrmFile); 
	File.append("ValidErodeFactor="+ValidErodeFactor, PrmFile); 
	File.append("nDefaultErode="+nDefaultErode, PrmFile); 
	File.append("MinSkeletonArea="+MinSkeletonArea, PrmFile); 
	File.append("MaxLocThkDisplayValue="+MaxLocThkDisplayValue, PrmFile); 	
	File.append("filterByAngle="+filterByAngle, PrmFile); 	
	File.append("minAngleToFilter="+minAngleToFilter, PrmFile); 	
	File.append("maxAngleToFilter="+maxAngleToFilter, PrmFile); 	
	File.append("ignoreBorderSize="+ignoreBorderSize, PrmFile); 
	File.append("ignoreTopDown="+ignoreTopDown, PrmFile); 
	File.append("ignoreFailedImagesFlag="+ignoreFailedImagesFlag, PrmFile); 
	File.append("BoundaryColor="+BoundaryColor, PrmFile); 
	File.append("CenterlineColor="+CenterlineColor, PrmFile); 
}


//===============================================================================================================
function setTimeString()
{
	MonthNames = newArray("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec");
	DayNames = newArray("Sun", "Mon","Tue","Wed","Thu","Fri","Sat");
	getDateAndTime(year, month, dayOfWeek, dayOfMonth, hour, minute, second, msec);
	TimeString ="Date: "+DayNames[dayOfWeek]+" ";
	if (dayOfMonth<10) {TimeString = TimeString+"0";}
	TimeString = TimeString+dayOfMonth+"-"+MonthNames[month]+"-"+year+", Time: ";
	if (hour<10) {TimeString = TimeString+"0";}
	TimeString = TimeString+hour+":";
	if (minute<10) {TimeString = TimeString+"0";}
	TimeString = TimeString+minute+":";
	if (second<10) {TimeString = TimeString+"0";}
	TimeString = TimeString+second;
}
  
//===============================================================================================================
function GenerateSummaryLines(SummaryTable)
{
	if (isOpen(SummaryTable))
	{
		IJ.renameResults(SummaryTable, "Results");
		Area = newArray(nResults);
		MeanLocThk = newArray(nResults);
		ModeLocThk = newArray(nResults);
		MedianLocThk = newArray(nResults);
		MinLocThk = newArray(nResults);
		MaxLocThk = newArray(nResults);
		for (n = 0; n < nResults ; n++)
		{
			Area[n] = getResult("Area", n); 
			MeanLocThk[n] = getResult("MeanLocThk", n); 
			ModeLocThk[n] = getResult("ModeLocThk", n); 
			MedianLocThk[n] = getResult("MedianLocThk", n); 
			MinLocThk[n] = getResult("MinLocThk", n); 
			MaxLocThk[n] = getResult("MaxLocThk", n); 
		}
		Array.getStatistics(Area, minArea, maxArea, meanArea, stdDevArea);
		Array.getStatistics(MeanLocThk, minMeanLocThk, maxMeanLocThk, meanMeanLocThk, stdDevMeanLocThk);
		Array.getStatistics(ModeLocThk, minModeLocThk, maxModeLocThk, meanModeLocThk, stdDevModeLocThk);
		Array.getStatistics(MedianLocThk, minMedianLocThk, maxMedianLocThk, meanMedianLocThk, stdDevMedianLocThk);
		Array.getStatistics(MinLocThk, minMinLocThk, maxMinLocThk, meanMinLocThk, stdDevMinLocThk);
		Array.getStatistics(MaxLocThk, minMaxLocThk, maxMaxLocThk, meanMaxLocThk, stdDevMaxLocThk);
		
		setResult("Label", nResults, "MeanValues"); 
		setResult("Area", nResults-1, meanArea); 
		setResult("MeanLocThk", nResults-1, meanMeanLocThk); 
		setResult("ModeLocThk", nResults-1, meanModeLocThk); 
		setResult("MedianLocThk", nResults-1, meanMedianLocThk); 
		setResult("MinLocThk", nResults-1, meanMinLocThk); 
		setResult("MaxLocThk", nResults-1, meanMaxLocThk); 
		
		setResult("Label", nResults, "StdValues"); 
		setResult("Area", nResults-1, stdDevArea); 
		setResult("MeanLocThk", nResults-1, stdDevMeanLocThk); 
		setResult("ModeLocThk", nResults-1, stdDevModeLocThk); 
		setResult("MedianLocThk", nResults-1, stdDevMedianLocThk); 
		setResult("MinLocThk", nResults-1, stdDevMinLocThk); 
		setResult("MaxLocThk", nResults-1, stdDevMaxLocThk); 

		setResult("Label", nResults, "MinValues"); 
		setResult("Area", nResults-1, minArea); 
		setResult("MeanLocThk", nResults-1, minMeanLocThk); 
		setResult("ModeLocThk", nResults-1, minModeLocThk); 
		setResult("MedianLocThk", nResults-1, minMedianLocThk); 
		setResult("MinLocThk", nResults-1, minMinLocThk); 
		setResult("MaxLocThk", nResults-1, minMaxLocThk); 

		setResult("Label", nResults, "MaxValues"); 
		setResult("Area", nResults-1, maxArea); 
		setResult("MeanLocThk", nResults-1, maxMeanLocThk); 
		setResult("ModeLocThk", nResults-1, maxModeLocThk); 
		setResult("MedianLocThk", nResults-1, maxMedianLocThk); 
		setResult("MinLocThk", nResults-1, maxMinLocThk); 
		setResult("MaxLocThk", nResults-1, maxMaxLocThk); 

		// Save Results - actual saving is done at the higher level function as this table include one line for each image
		IJ.renameResults("Results", SummaryTable); // rename to avoid table overwrite
				
	}
}


