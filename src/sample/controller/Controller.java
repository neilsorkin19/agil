package sample.controller;

// import local things

import javafx.application.Platform;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.MenuItem;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import org.opencv.core.Mat;
import org.opencv.videoio.VideoCapture;

import sample.Main;
import sample.model.StudentResponseMetadata;
import sample.util.ExcelManager;
import sample.util.ImageProcessor;
import sample.util.Utils;

import javax.imageio.ImageIO;
import javax.print.PrintService;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.JobName;
import javax.print.attribute.standard.OrientationRequested;
import javax.sound.midi.*;
import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.awt.print.PageFormat;
import java.awt.print.Printable;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

//javafx imports
//opencv imports


public class Controller {
    //to determine if the key is being scanned at the moment
    @FXML
    private CheckBox scanKey;
    // the FXML button
    @FXML
    private Button startCam;
    // the FXML image view
    @FXML
    private CheckBox soundEnabled;
    @FXML
    private TextField output;
    @FXML
    private ImageView currentFrame;
    @FXML
    private ImageView originalImage;
    @FXML
    private TableView responseTable;
    @FXML
    private TableColumn<String, StudentResponseMetadata> versionCol;
    @FXML
    private TableColumn<String, StudentResponseMetadata> correctCol;
    @FXML
    private TableColumn<String, StudentResponseMetadata> percentCol;
    @FXML
    private TableColumn<String, StudentResponseMetadata> idCol;
    @FXML
    private MenuButton dropDownPrintMenu;
    @FXML
    private MenuItem twoSheets;
    @FXML
    private MenuItem oneSheet;

    // a timer for acquiring the video stream
    private ScheduledExecutorService timer;
    // the OpenCV object that realizes the video capture
    private VideoCapture capture = new VideoCapture();
    // a flag to change the button behavior
    private boolean cameraActive = false;
    // the id of the camera to be used
    private static int cameraId = 0; // first camera

    private ImageProcessor imageProcessor;
    private String[][] keyVals = new String[5][60]; // will be entered by teacher
    private int requiredConsecutiveMatchingAnswers = 3;
    public static int keyVersion = 0; // default to key version

    public static ExcelManager excelManager = new ExcelManager();
    public static boolean hasData = false;
    private String lastId = "";
    public static boolean isWarning = false; // to show whether or not the user has pressed ok on the warning

    //the behavior of the selected key is if the version is not selected, it will be the same as the previous person's key version

    public void setUpResponseTable() {
        if (responseTable.getSelectionModel().getSelectionMode() != SelectionMode.MULTIPLE) {
            versionCol.setCellValueFactory(new PropertyValueFactory<>("keyVersion"));
            correctCol.setCellValueFactory(new PropertyValueFactory<>("numCorrect"));
            percentCol.setCellValueFactory(new PropertyValueFactory<>("percentCorrect"));
            idCol.setCellValueFactory(new PropertyValueFactory<>("id"));
            MultipleSelectionModel multipleSelectionModel = responseTable.getSelectionModel(); // multiple selections allowed
            multipleSelectionModel.setSelectionMode(SelectionMode.MULTIPLE);
            responseTable.setOnKeyPressed(new EventHandler<KeyEvent>() {
                @Override
                public void handle(KeyEvent keyEvent) {
                    final ObservableList selectedItems = responseTable.getSelectionModel().getSelectedItems();
                    if (selectedItems != null) {
                        if (keyEvent.getCode().equals(KeyCode.DELETE)) {
                            responseTable.getItems().removeAll(selectedItems);
                        }
                    }
                }
            });
        }
    }

    @FXML
    protected void createNew(ActionEvent event) {
        File emptyFile = new File(FileSystemView.getFileSystemView().getDefaultDirectory().getPath());
        FileChooser fileChooser = new FileChooser();
        fileChooser.setInitialDirectory(emptyFile);
        fileChooser.setTitle("Select Excel File Location");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter(
                "Excel File (*.xls, *.xlsx)",
                "*.xls", "*.xlsx"));
        fileChooser.setInitialFileName("*.xlsx");
        File chosenFile = fileChooser.showSaveDialog(new Stage());
        if (chosenFile != null) {
            if (chosenFile.getName().endsWith(".xlsx") || chosenFile.getName().endsWith(".xls")) { // good naming
                excelManager.createNewAnswerBook(chosenFile);
                output.setText(chosenFile.getName() + " created and is current workbook");
                setUpResponseTable();
                responseTable.getItems().removeAll(responseTable.getItems()); // delete everything
            } else { // bad extension
                output.setText("Bad file extension for: " + chosenFile.getName());
            }
        } else { // no selection (cancel)
            output.setText("Operation cancelled");
        }
    }

    @FXML
    protected void addToExisting(ActionEvent event) {
        File emptyFile = new File(FileSystemView.getFileSystemView().getDefaultDirectory().getPath());
        FileChooser fileChooser = new FileChooser();
        fileChooser.setInitialDirectory(emptyFile);
        fileChooser.setTitle("Open Excel File");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter(
                "Excel File (*.xls, *.xlsx)",
                "*.xls", "*.xlsx"));
        File chosenFile = fileChooser.showOpenDialog(new Stage());
        if (chosenFile != null) {
            if (chosenFile.getName().endsWith(".xlsx") || chosenFile.getName().endsWith(".xls")) { // good naming
                String outputVal = excelManager.addToExistingBook(chosenFile);
                output.setText(outputVal);
                if (excelManager.isAssigned()) {
                    this.keyVals = excelManager.alterKeysBasedOnKeysFromSheet(keyVals.clone());
                    setUpResponseTable();
                    responseTable.getItems().clear();
                    ArrayList<StudentResponseMetadata>[] allStudentScoreMetaData = excelManager.getAllStudentMetadataInSheet();
                    for (ArrayList<StudentResponseMetadata> metaList : allStudentScoreMetaData) {
                        if (metaList != null)
                            for (StudentResponseMetadata metaStudent : metaList) {
                                System.out.println(metaStudent.getId());
                                StudentResponseMetadata metaCopy = new StudentResponseMetadata(
                                        metaStudent.getKeyVersion(),
                                        metaStudent.getNumCorrect(),
                                        metaStudent.getPercentCorrect(),
                                        metaStudent.getId()
                                );
                                responseTable.getItems().add(metaCopy);
                            }
                    }
                }
            } else { // bad extension
                output.setText("Bad file extension for: " + chosenFile.getName());
            }
        } else { // no selection (cancel)
            output.setText("Selection cancelled");
        }
    }

    /**
     * The action triggered by pushing the button on the GUI
     *
     * @param event the push button event
     */
    @FXML
    protected void startCamera(ActionEvent event) {
        imageProcessor = new ImageProcessor();
        setUpResponseTable();
        if (!this.cameraActive) {
            // start the video capture
            this.capture.open(cameraId);
            // is the video stream available?
            if (this.capture.isOpened()) {
                this.cameraActive = true;

                int highValue = 10000;
                capture.set(3, highValue);
                capture.set(4, highValue);

                ArrayList<String[]> lastThreeResponses = new ArrayList<>();


                // grab a frame every 33 ms (30 frames/sec)
                Runnable frameGrabber = new Runnable() {

                    @Override
                    public void run() {
                        currentFrame.setFitWidth(currentFrame.getParent().getScene().getWidth());

                        // effectively grab and process a single frame
                        Mat frame = grabFrame();
                        Mat origFrame = new Mat();
                        frame.copyTo(origFrame);

                        Image origImageToShow = Utils.mat2Image(origFrame);
                        updateImageView(originalImage, origImageToShow);

                        String[] studentResponses = new String[60]; // current frame responses
                        frame = imageProcessor.process(frame, studentResponses); // process the image
                        //System.out.println(Controller.keyVersion);

                        if (lastThreeResponses.size() >= requiredConsecutiveMatchingAnswers) {
                            ArrayList<String[]> newThree = new ArrayList<>();
                            newThree.add(studentResponses);
                            newThree.add(lastThreeResponses.get(0));
                            newThree.add(lastThreeResponses.get(1));
                            lastThreeResponses.set(0, newThree.get(0));
                            lastThreeResponses.set(1, newThree.get(1));
                            lastThreeResponses.set(2, newThree.get(2));
                        } else {
                            lastThreeResponses.add(0, studentResponses);
                        }

                        // if three consecutive matches occur, it is probably good, so use it
                        //https://stackoverflow.com/questions/16462854/midi-beginner-need-to-play-one-note
                        if (lastThreeResponses.size() == requiredConsecutiveMatchingAnswers &&
                                lastThreeResponses.get(2)[0] != "O" && // not blank matches
                                Arrays.equals(lastThreeResponses.get(0), lastThreeResponses.get(1)) &&
                                Arrays.equals(lastThreeResponses.get(1), lastThreeResponses.get(2))) {
                            try {
                                lastThreeResponses.clear();

                                try {
                                    /* Create a new Sythesizer and open it. Most of
                                     * the methods you will want to use to expand on this
                                     * example can be found in the Java documentation here:
                                     * https://docs.oracle.com/javase/7/docs/api/javax/sound/midi/Synthesizer.html
                                     */

                                    Synthesizer midiSynth = MidiSystem.getSynthesizer();
                                    midiSynth.open();

                                    //get and load default instrument and channel lists
                                    Instrument[] instr = midiSynth.getDefaultSoundbank().getInstruments();
                                    MidiChannel[] mChannels = midiSynth.getChannels();

                                    midiSynth.loadInstrument(instr[0]);//load an instrument
                                    if (!isWarning) {
                                        if (soundEnabled.isSelected()) {
                                            mChannels[0].noteOn(80, 60);//On channel 0, play note number 80 with velocity 100
                                        }
                                        String convertedID = convertFromSheetToFiveDigitID(studentResponses);

                                        if (scanKey.isSelected()) { // key mode
                                            if (!convertedID.equals("N/A") && // the response has a non-blank ID
                                                    !lastId.equals(convertedID)) { // and it is not the same as the last
                                                Platform.runLater(new Runnable() {
                                                    @Override
                                                    public void run() {
                                                        isWarning = true;
                                                        Main.promptWarningKeyHasId();
                                                        isWarning = false;
                                                    }
                                                });
                                            } else {
                                                keyVals[keyVersion] = studentResponses; // store the key
                                                hasData = true; // make it known that there is some data here
                                                if (excelManager.isAssigned()) { // they have added a spreadsheet
                                                    excelManager.addKeyToBook(convertFromNumberToKeyID(keyVersion), studentResponses); // add to excel
                                                }
                                            }
                                        } else {
                                            try {
                                                grade(keyVals[keyVersion], studentResponses); // grade and write the percentage to the screen
                                            } catch (NullPointerException ne) { // the key is not scanned for that version
                                                ne.printStackTrace();
                                                if (soundEnabled.isSelected())
                                                    mChannels[1].noteOn(67, 80);//On channel 1, play note number 67 with velocity 100
                                                output.setText("Key " + convertFromNumberToKeyID(keyVersion) + " has not been scanned yet.");
                                            }
                                        }
                                        lastId = convertedID;
                                    }
                                } catch (MidiUnavailableException e) {
                                    System.err.println("No Midi available");
                                }

                            } catch (Exception e) {
                                e.printStackTrace();
                                System.err.println("Waited too long. " + e.toString());
                            }
                        }

                        // convert and show the frame
                        Image imageToShow = Utils.mat2Image(frame);//update the view on the processed frame
                        updateImageView(currentFrame, imageToShow);

                        origFrame.release();
                        frame.release();
                    }
                };

                this.timer = Executors.newSingleThreadScheduledExecutor();
                this.timer.scheduleAtFixedRate(frameGrabber, 200, 33, TimeUnit.MILLISECONDS); // 200 ms delay for warmup

                // update the button content
                this.startCam.setText("Stop Camera");
            } else {
                // log the error
                System.err.println("Impossible to open the camera connection...");
            }
        } else {
            // the camera is not active at this point
            this.cameraActive = false;
            // update again the button content
            this.startCam.setText("Start Camera");

            // stop the timer
            this.stopAcquisition();
        }
    }

    private void grade(String[] keyVals, String[] studentVals) {
        double numCorrect = 0;
        double numQuestions = 0;
        for (int question = 0; question < keyVals.length; question++) {
            if (keyVals[question].equals("O")) { // blank answer in the key means the test is over
                break;
            }
            if (keyVals[question].equals(studentVals[question])) { // matching answers
                numCorrect++;
            }
            numQuestions++;
        }
        if (numQuestions > 0) { // div by 0 error catching
            double percentCorrect = numCorrect / numQuestions;
            String writeToScreen = numCorrect + " / " + numQuestions + " = " + (100 * percentCorrect) + "%";
            String fiveDigitID = convertFromSheetToFiveDigitID(studentVals);
            output.setText(writeToScreen + ", ID: " + fiveDigitID);

            StudentResponseMetadata studentResponseMetadata = new StudentResponseMetadata(
                    convertFromNumberToKeyID(keyVersion),
                    (int) numCorrect,
                    percentCorrect * 100,
                    fiveDigitID
            );
            int indexOfMatchingID = -1;
            for (int i = 0; i < responseTable.getItems().size(); i++) {
                StudentResponseMetadata currentStudent = (StudentResponseMetadata) responseTable.getItems().get(i);
                if (currentStudent.getId().equals(studentResponseMetadata.getId())) {
                    indexOfMatchingID = i;
                    responseTable.getItems().set(i, studentResponseMetadata); // set the second one to overwrite
                }
            }
            if (indexOfMatchingID == -1) {
                responseTable.getItems().add(studentResponseMetadata); // output to table
            }
            if (excelManager.isAssigned()) { // they have set up a book
                excelManager.addStudentResponsesToBook(
                        convertFromNumberToKeyID(keyVersion),
                        convertFromSheetToFiveDigitID(studentVals),
                        studentVals); // add to excel
            }
        }
    }

    private String convertFromNumberToKeyID(int id) {
        switch (id) {
            case 0:
                return "A";
            case 1:
                return "B";
            case 2:
                return "C";
            case 3:
                return "D";
            case 4:
                return "E";
            default:
                return "unknown";
        }
    }

    private String convertFromSheetToFiveDigitID(String[] studentVals) {
        String[] options = {"A", "B", "C", "D", "E"}; // all possible values
        ArrayList<String> idList = new ArrayList<>(); //accumulates the number in order
        for (int i = 0; i < 5; i++) {
            idList.add("-1");
        }
        boolean leftBlank = true;
        for (int question = 50; question < 60; question++) {
            String studentRespForGivenQuestion = studentVals[question];
            for (int possibleChoice = 0; possibleChoice < options.length; possibleChoice++) {
                if (!studentRespForGivenQuestion.equals("O") &&
                        studentRespForGivenQuestion.contains(options[possibleChoice])) {
                    idList.set(possibleChoice, Integer.toString((question - 49) % 10));
                    leftBlank = false;
                }
            }
        }
        idList.removeIf(value -> (value.equals("-1")));
        String compiledVals = "";
        for (String val : idList) {
            compiledVals += val;
        }
        if (leftBlank) {
            return "N/A";
        }
        return compiledVals;
    }

    /**
     * Get a frame from the opened video stream (if any)
     *
     * @return the {@link Mat} to show
     */
    private Mat grabFrame() {
        // init everything
        Mat frame = new Mat();

        // check if the capture is open
        if (this.capture.isOpened()) {
            try {
                // read the current frame
                this.capture.read(frame);
            } catch (Exception e) {
                // log the error
                System.err.println("Exception during the image elaboration: " + e);
            }
        }
        return frame;
    }

    /**
     * Stop the acquisition from the camera and release all the resources
     */
    private void stopAcquisition() {
        if (this.timer != null && !this.timer.isShutdown()) {
            try {
                // stop the timer
                this.timer.shutdown();
                this.timer.awaitTermination(33, TimeUnit.MILLISECONDS);
            } catch (InterruptedException e) {
                // log any exception
                System.err.println("Exception in stopping the frame capture, trying to release the camera now... " + e);
            }
        }

        if (this.capture.isOpened()) {
            // release the camera
            this.capture.release();
        }
    }

    /**
     * Update the {@link ImageView} in the JavaFX main thread
     *
     * @param view  the {@link ImageView} to update
     * @param image the {@link Image} to show
     */
    private void updateImageView(ImageView view, Image image) {
        Utils.onFXThread(view.imageProperty(), image);
    }

    /**
     * On application close, stop the acquisition from the camera
     */
    protected void setClosed() {
        this.stopAcquisition();
    }

    private ExecutorService executorService = Executors.newCachedThreadPool(); // for prints

    @FXML
    private void printTwoSheets() {
        executorService.execute(new Runnable() {
            @Override
            public void run() {
                try {
                    BufferedImage bufferedImage = ImageIO.read(Main.class.getResource("images/FormTemplateWithID_2.png"));
                    printImage(bufferedImage, 2);
                } catch (IOException ioe) {
                    ioe.printStackTrace();
                }
            }
        });
    }

    private static BufferedImage rotateClockwise90(BufferedImage src) {
        int width = src.getWidth();
        int height = src.getHeight();

        BufferedImage dest = new BufferedImage(height, width, src.getType());

        Graphics2D graphics2D = dest.createGraphics();
        graphics2D.translate((height - width) / 2, (height - width) / 2);
        graphics2D.rotate(Math.PI / 2, (float) width / 2, (float) height / 2);
        graphics2D.drawRenderedImage(src, null);
        return dest;
    }

    @FXML
    private void printOneSheet() {
        executorService.execute(new Runnable() {
            @Override
            public void run() {
                try {
                    BufferedImage bufferedImage = ImageIO.read(Main.class.getResource("images/FormTemplateWithID.png"));
                    printImage(bufferedImage, 1);
                } catch (IOException ioe) {
                    ioe.printStackTrace();
                }
            }
        });
    }

    private void printImage(BufferedImage image, int sheetsPerPage) {
        PrinterJob printerJob = PrinterJob.getPrinterJob();
        PageFormat pf = printerJob.defaultPage();
        PrintRequestAttributeSet aset = new HashPrintRequestAttributeSet();
        PageFormat validatePage = printerJob.validatePage(pf);
        printerJob.setPrintable((graphics, pageFormat, pageIndex) -> {
            if (pageIndex > 0) {
                return Printable.NO_SUCH_PAGE;
            }
            graphics.translate((int) pageFormat.getImageableX(), (int) pageFormat.getImageableY());
            double width = pageFormat.getImageableWidth();
            double height = pageFormat.getImageableHeight();
            graphics.drawImage(image, 0, 0, (int) width, (int) height, null);
            return Printable.PAGE_EXISTS;
        }, validatePage);
        aset.add(new JobName("Agil Grader Sheets", null));
        if(sheetsPerPage == 2) // if 2 50Q sheets, make landscape
            aset.add(OrientationRequested.LANDSCAPE);
        try {
            PrintService[] services = PrinterJob.lookupPrintServices();
            printerJob.setJobName("Agil Grader Sheets");
            if (services.length > 0) {
                if (printerJob.printDialog(aset)) {
                    printerJob.print(aset);
                    output.setText("Printing " + printerJob.getCopies() * sheetsPerPage + " 50Q sheet(s)");
                }
                else{
                    output.setText("Print cancelled");
                }
            } else {
                output.setText("No printers detected");
            }
        } catch (PrinterException pe) {
            output.setText("Error printing");
            pe.printStackTrace();
        }
    }
}
