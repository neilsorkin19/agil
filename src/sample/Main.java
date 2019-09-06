package sample;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.scene.layout.BorderPane;
import javafx.stage.Stage;

import org.opencv.core.Core;

import sample.controller.Controller;

import java.io.IOException;
import java.util.Optional;

public class Main extends Application {
    private Stage primaryStage;
    @Override
    public void start(Stage primaryStage) {
        this.primaryStage = primaryStage;
        this.primaryStage.setTitle("Agil");
        initRootLayout();
    }

    // Compulsory
    static {
        
        try {
            System.loadLibrary(Core.NATIVE_LIBRARY_NAME);
        } catch (Exception e) {
            System.out.println(e.toString());
        }
    }
    public static void main(String[] args) {
        launch(args);
    }

    private void initRootLayout() {
        try{
            FXMLLoader loader = new FXMLLoader();
            loader.setLocation(Main.class.getResource("view/sample.fxml"));
            BorderPane rootLayout = loader.load();

            Scene scene = new Scene(rootLayout, 1080, 720);
            primaryStage.setScene(scene);
            primaryStage.setMaximized(true);
            //https://stackoverflow.com/questions/12153622/how-to-close-a-javafx-application-on-window-close
            primaryStage.setOnCloseRequest(t -> {
                if(Controller.excelManager.isAssigned() && Controller.hasData) {
                    if(!Controller.excelManager.canSave()){
                        promptConfirm(); // make sure they want to close without saving
                    }
                    else {
                        Controller.excelManager.saveBook();
                        forceClose();
                    }
                }
                else {
                    if(Controller.hasData) // there is data
                        promptConfirm();
                    else
                        forceClose();
                }
            });
            primaryStage.show();
        }
        catch (IOException ioe){
            ioe.printStackTrace();
        }
    }

    public static void forceClose(){
        if(Controller.excelManager.isAssigned()){
            Controller.excelManager.closeBook();
        }
        Platform.exit();
        System.exit(0);
    }

    public static void promptConfirm(){
        Alert confirmationClose = new Alert(
                Alert.AlertType.CONFIRMATION);
        confirmationClose.setTitle("Close without saving");
        confirmationClose.setHeaderText("Close without saving responses?");
        confirmationClose.setContentText("There is no file set up properly to save responses to. " +
                "This may be due to the file being open, or never selected.");
        Optional<ButtonType> optionChosen = confirmationClose.showAndWait();
        if(optionChosen.isPresent()) {
            if (optionChosen.get() == ButtonType.OK) {// ok selection
                forceClose();
            }
            else if(optionChosen.get() == ButtonType.CANCEL){ // don't close

            }
            else if(optionChosen.get() == ButtonType.CLOSE){ // closed the warning window, treat as cancel

            }
        }
    }
    public static void promptWarningKeyHasId(){
        Alert confirmAllowingAKeyWithIDToBeKey = new Alert(
                Alert.AlertType.WARNING,
                "This key has an ID on it, meaning it is most likely a student response.");
        confirmAllowingAKeyWithIDToBeKey.setTitle("Continue?");
        confirmAllowingAKeyWithIDToBeKey.setHeaderText("Potentially scanning student sheet as key");
        confirmAllowingAKeyWithIDToBeKey.setContentText("There is a valid ID in the 5-Digit ID section of the sheet you are scanning. " +
                "This is generally unintended, as student response sheets are the only sheets that have the 5-Digit ID section filled in. " +
                "To scan as key anyways, select ok and scan again.");
        confirmAllowingAKeyWithIDToBeKey.showAndWait();
    }
    public static void promptWarningOpenFile(){
        Alert confirmAllowingAKeyWithIDToBeKey = new Alert(
                Alert.AlertType.WARNING,
                "The file you selected is open. Close it, then try again.");
        confirmAllowingAKeyWithIDToBeKey.showAndWait();
    }
}
