package sample.util;

import java.io.File;
import java.net.URISyntaxException;
import org.opencv.calib3d.Calib3d;
import org.opencv.core.*;
import org.opencv.features2d.DescriptorMatcher;
import org.opencv.features2d.Features2d;
import org.opencv.features2d.FlannBasedMatcher;
import org.opencv.features2d.ORB;
import org.opencv.imgcodecs.Imgcodecs;
import org.opencv.imgproc.Imgproc;

import sample.controller.Controller;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;
import sample.Main;


public class ImageProcessor {

    private static MatOfKeyPoint mkObj;
    private static Mat dsObj;
    private static Mat templateImg; //read form as color image
    private static Mat templateClone;

    public ImageProcessor() {
        mkObj = new MatOfKeyPoint();
        dsObj = new Mat();
        templateImg = templateImage();
        templateClone = templateImg.clone();
        Imgproc.resize(templateClone, templateClone, new Size(templateClone.cols() / 3, templateClone.rows() / 3), 0, 0, Imgproc.INTER_LINEAR_EXACT);
        Imgproc.cvtColor(templateClone, templateClone, Imgproc.COLOR_BGR2GRAY);
        ORB orbDetectorTemplate = ORB.create(2000, 1.2f, 5, 0, 0, 2, ORB.HARRIS_SCORE, 80);
        orbDetectorTemplate.detectAndCompute(templateClone, new Mat(), mkObj, dsObj);
    }


    //https://stackoverflow.com/questions/15719977/using-opencv-java-bindings-to-read-an-image
    //https://www.learnopencv.com/homography-examples-using-opencv-python-c/

    public Mat process(Mat imgMat, String[] studentResponses) {
        try {
            Mat img = imgMat.clone();
            Arrays.fill(studentResponses, "O"); // blank vals to fill
            Mat orbMat = orbProc(img);
            if (orbMat.cols() < img.cols()) {// good enough match because the filtered sheet width is smaller than the original
                Mat templateClone = templateImg.clone();
                Imgproc.resize(templateImg, templateClone, new Size(templateClone.cols() / 3, templateClone.rows() / 3), 0, 0, Imgproc.INTER_LINEAR_EXACT);
                orbMat = binarize(templateClone, orbMat, studentResponses);
                img = orbMat;
            }
            return img;
        }
        catch (Exception e){
            System.out.println(e.toString());
        }
        return imgMat;
    }

    private static Mat templateImage() {
        final String path = System.getProperty("user.dir") + "/images/FormTemplateWithID.png";
        System.out.println(path);
        return Imgcodecs.imread(path);
    }

    private static Mat orbProc(Mat src) {
        Mat imgClone = src.clone();
        Mat colorSrc = src.clone();
        //greyscale
        Imgproc.cvtColor(imgClone, imgClone, Imgproc.COLOR_BGR2GRAY);
        ORB orbDetector = ORB.create(800, 1.2f, 2, 0, 0, 2, ORB.HARRIS_SCORE, 80);

        MatOfKeyPoint mkScene = new MatOfKeyPoint();
        Mat dsScene = new Mat();

        orbDetector.detectAndCompute(imgClone, new Mat(), mkScene, dsScene);

        List<MatOfDMatch> knnMatches = new ArrayList<>();

        //FLANN Based Matcher
        DescriptorMatcher f = FlannBasedMatcher.create(FlannBasedMatcher.BRUTEFORCE_HAMMING);
        f.knnMatch(dsObj, dsScene, knnMatches, 2);

        //-- Filter matches using the Lowe's ratio test
        float ratioThresh = 0.85f;
        LinkedList<DMatch> listOfGoodMatches = new LinkedList<>();

        for (int i = 0; i < knnMatches.size(); i++) {
            if (knnMatches.get(i).rows() > 1) {
                DMatch[] matches = knnMatches.get(i).toArray();
                if (matches[0].distance < ratioThresh * matches[1].distance) {
                    listOfGoodMatches.add(matches[0]);
                    //System.out.println(matches[0]);
                }
            }
        }
        int numOfBestMatches = 40;
        List<DMatch> bestMatches = new LinkedList<>();
        for (int m = 0; m < numOfBestMatches && m < listOfGoodMatches.size(); m++) {//add first numOfBestMatches
            bestMatches.add(listOfGoodMatches.get(m));
        }
        if (bestMatches.size() == numOfBestMatches) {
            for (int k = numOfBestMatches; k < listOfGoodMatches.size(); k++) {//cycle all good matches
                for (int match = 0; match < numOfBestMatches; match++) {//compare them to the best numOfBestMatches
                    if (bestMatches.get(match).distance > listOfGoodMatches.get(k).distance) {
                        bestMatches.add(match, listOfGoodMatches.get(k));//add it to the good list at the proper index
                        match = numOfBestMatches;//break for loop
                    }
                }
                //truncate anything larger than numOfBestMatches
                bestMatches = bestMatches.stream().limit(numOfBestMatches).collect(Collectors.toList());
            }
        }

        MatOfDMatch goodMatches = new MatOfDMatch();
        goodMatches.fromList(bestMatches);
        Features2d.drawMatches(templateClone, mkObj, colorSrc, mkScene, goodMatches, src, new Scalar(100, 0, 0), new Scalar(255, 0, 255), new MatOfByte(), Features2d.DrawMatchesFlags_DEFAULT);
        if (goodMatches.size().height >= numOfBestMatches) {//enough good matches
            //Features2d.drawKeypoints(colorSrc, mkScene, colorSrc, new Scalar(255, 50, 80), Features2d.DrawMatchesFlags_NOT_DRAW_SINGLE_POINTS);
            //https://docs.opencv.org/3.4/d7/dff/tutorial_feature_homography.html
            //-- Localize the object
            List<Point> obj = new ArrayList<>();
            List<Point> scene = new ArrayList<>();
            List<KeyPoint> listOfKeypointsObject = mkObj.toList();
            List<KeyPoint> listOfKeypointsScene = mkScene.toList();
            for (int i = 0; i < listOfGoodMatches.size(); i++) {
                //-- Get the keypoints from the good matches
                obj.add(listOfKeypointsObject.get(listOfGoodMatches.get(i).queryIdx).pt);
                scene.add(listOfKeypointsScene.get(listOfGoodMatches.get(i).trainIdx).pt);
            }
            MatOfPoint2f objMat = new MatOfPoint2f(), sceneMat = new MatOfPoint2f();
            objMat.fromList(obj);
            sceneMat.fromList(scene);
            double ransacReprojThreshold = 1.0;
            Mat H = Calib3d.findHomography(objMat, sceneMat, Calib3d.RANSAC, ransacReprojThreshold);
            Imgproc.warpPerspective(colorSrc, colorSrc, H.inv(), templateClone.size());
            if (isAligned(colorSrc, H.inv())) {
                return colorSrc;
            } else {
                return src;
            }
            //Features2d.drawMatches(templateClone, mkObj, colorSrc, mkScene, goodMatches, src, new Scalar(100, 0, 0), new Scalar(255, 0, 255), new MatOfByte(), Features2d.DrawMatchesFlags_NOT_DRAW_SINGLE_POINTS);
        }
        return src;
    }

    //pre: homo is a Mat with a homography transformation applied to it
    //post: true if the Mat is aligned enough; false if no aligned enough
    private static boolean isAligned(Mat homo, Mat warp) {
        //calculate the determinant
        Mat discMat = warp.submat(0, 1, 0, 1);
        double determinant = Core.determinant(discMat);

        //count the non-zero pixels
        int nonZero = 0;
        for (int row = 0; row < homo.rows(); row++) {
            for (int col = 0; col < homo.cols(); col++) {
                double[] color = homo.get(row, col); //gets color at row and col in double[] format
                if (color[0] == 0 && color[0] == color[1] && color[0] == color[2]) {
                }//if all values are zero, do nothing
                else {//otherwise, increment nonZero
                    nonZero++;
                }
            }
        }

        int pixelsInHomo = homo.rows() * homo.cols();
        double percentNonZero = nonZero / pixelsInHomo; // must be 99% non-zero pixels
        double percentNonZeroThreshold = 0.99;

        //determine if the image is a good fit
        if (determinant < 1.1) {
            return percentNonZero > percentNonZeroThreshold;
        }
        return false;
    }

    //sheet is oriented at this point
    private static Mat binarize(Mat template, Mat studentSheet, String[] studentResponses) {
        Mat colorStudentSheet = studentSheet.clone();
        Mat grayStudentSheet = new Mat();
        Imgproc.cvtColor(colorStudentSheet, grayStudentSheet, Imgproc.COLOR_BGR2GRAY);

        Mat grayTemplate = template.clone();
        Mat edges = new Mat();
        Imgproc.cvtColor(grayTemplate, grayTemplate, Imgproc.COLOR_BGR2GRAY);//convert to gray
        Imgproc.blur(grayTemplate, grayTemplate, new Size(5, 5)); //blur image
        Imgproc.Canny(grayTemplate, edges, 10, 15);//canny edge
        Mat circles = new Mat();

        int minR = grayTemplate.cols() / 64;
        int maxR = grayTemplate.cols() / 50;

        Imgproc.HoughCircles(edges, circles, Imgproc.CV_HOUGH_GRADIENT, 1.00, maxR * 2, 10, 14, minR, maxR); // find circles
        Imgproc.cvtColor(edges, edges, Imgproc.COLOR_GRAY2BGR); //convert to color


        if (circles.cols() > 0) {
            ArrayList<double[]> allHits = new ArrayList<>();

            for (int x = 0; x < Math.max(circles.cols(), circles.rows()); x++) { // drawing and detecting
                double[] circleVec = circles.get(0, x);

                if (circleVec == null) {
                    break;
                }

                Point center = new Point((int) circleVec[0], (int) circleVec[1]);

                int radius = (int) circleVec[2]; // needed for surrounding darkness comparison calculation

                double[] surroundingLowerRight = grayStudentSheet.get((int) center.y + radius, (int) center.x + radius); //down right
                double[] surroundingUpperRight = grayStudentSheet.get((int) center.y - radius, (int) center.x + radius); // up right
                double[] surroundingUpperLeft = grayStudentSheet.get((int) center.y - radius, (int) center.x - radius); // up left
                double[] surroundingLowerLeft = grayStudentSheet.get((int) center.y + radius, (int) center.x - radius); //down left

                double[] colorAtPointCenter = grayStudentSheet.get((int) center.y, (int) center.x);
                //take 4 more diagonal samples just in case
                double[] colorAtPoint2 = grayStudentSheet.get((int) center.y + 1, (int) center.x + 1); //down right
                double[] colorAtPoint3 = grayStudentSheet.get((int) center.y - 1, (int) center.x + 1); // up right
                double[] colorAtPoint4 = grayStudentSheet.get((int) center.y - 1, (int) center.x - 1); // up left
                double[] colorAtPoint5 = grayStudentSheet.get((int) center.y + 1, (int) center.x - 1); //down left

                double avgSurroundingIntensity = (surroundingLowerRight[0] + surroundingUpperRight[0] + surroundingUpperLeft[0] + surroundingLowerLeft[0]) / 4;
                double avgIntensity = (colorAtPointCenter[0] + colorAtPoint2[0] + colorAtPoint3[0] + colorAtPoint4[0] + colorAtPoint5[0]) / 5; // average intensity of the 5 points

                if (avgIntensity < avgSurroundingIntensity * 0.85) { // must be a factor less intense than normal
                    Imgproc.circle(colorStudentSheet, center, 3, new Scalar(0, 255, 0), 2);
                    allHits.add(circleVec); // adds all hits to the list to be sorted later into a grid
                } else
                    Imgproc.circle(colorStudentSheet, center, 3, new Scalar(0, 0, 255), 3);
            }

            allHits.sort((o1, o2) -> (int) (o1[0] - o2[0]));// sort based on x
            // pick first value (key id); a = 191, b = 243, c = 294, d = 345, e = 396; 51 is the avg dist between the x vals
            double yKey = allHits.get(0)[1]; // get the y value of the key version
            double[] keyYValues = {191, 243, 294, 345, 396};
            String[] selectionVals = {"A", "B", "C", "D", "E"};
            for (int key = 0; key < keyYValues.length; key++) {
                if (Math.abs(yKey - keyYValues[key]) < 5) { // first key close enough in distance to selected vals

                    Controller.keyVersion = key; // store the key version
                    //System.out.println("KEY: " + selectionVals[key]);
                    allHits.removeIf(bubble -> (bubble[0] < 100)); // removes all key values (in case more than one was selected for some reason)
                    break;
                }
            }

            allHits.sort((o1, o2) -> (int) (o1[1] - o2[1]));// sort based on y

            // 20 cols of selection possibilities
            double[] selectionCols = {
                    137, 172, 206, 242, 275, //1 - 20
                    359, 393, 428, 463, 498, // 21 - 40
                    587, 614, 647, 684, 718}; // 41 - 60

            // 20 rows of selection possibilities
            double[] selectionRows = {
                    194.5, 234.5, 275.5, 316.5, 357.5, 397.5, 439.5, 480.5, 521.5, 560.5, // 1 - 10
                    642.5, 683.5, 724.5, 765.5, 805.5, 846.5, 884.5, 928.5, 969.5, 1008.5}; // 11 - 20


            for (double[] bubble : allHits) {
                for (int r = 0; r < selectionRows.length; r++) {
                    for (int c = 0; c < selectionCols.length; c++) {
                        if ((Math.abs(bubble[0] - selectionCols[c]) < 15) && (Math.abs(bubble[1] - selectionRows[r]) < 15)) { // very close to a predefined row and col
                            if (studentResponses[((r + 1) + (20 * (c / 5))) - 1].equals("O")) // blank case
                                studentResponses[((r + 1) + (20 * (c / 5))) - 1] = selectionVals[c % 5]; // set to new val
                            else
                                studentResponses[((r + 1) + (20 * (c / 5))) - 1] += selectionVals[c % 5]; // append letter to other letters that might exist
                            //System.out.println("Question: " + ((r + 1) + (20 * (c / 5))) + " Ans: " + studentResponses[((r + 1) + (20 * (c / 5))) - 1]);
                            break;
                        }
                    }
                }
                //System.out.println("x: " + bubble[0] + " y: " + bubble[1]);
            }
            /*
            for (int question = 0; question < studentResponses.length; question++) {
                if (studentResponses[question] != null)
                    System.out.println("Question: " + (question + 1) + " Ans: " + studentResponses[question]);
            }
             */

        }

        return colorStudentSheet;
    }


    private static void hueFilter(Mat src) {
        Imgproc.cvtColor(src, src, Imgproc.COLOR_BGR2HSV);//convert color space to HSV
        Mat srcClone = src.clone();
        Scalar colorMultiplier = new Scalar(0, 0, 0); //black for mask
        Mat mask = new Mat();
        //Scalar.all(0)
        Mat zeros = new Mat(srcClone.rows(), srcClone.cols(), CvType.CV_8UC3, colorMultiplier);//fill whole mat with black

        //Core.inRange(srcClone, lower, upper, mask); //create a mask with values in range
        Core.bitwise_not(mask, mask);//apply mask
        Core.bitwise_and(srcClone, zeros, src, mask);//only keep values in mask

        Imgproc.cvtColor(src, src, Imgproc.COLOR_HSV2BGR);//convert color back to BGR
    }

}