package sample.model;

public class StudentResponseMetadata {
    private String keyVersion;
    private Integer numCorrect;
    private Double percentCorrect;
    private String id;


    public StudentResponseMetadata(){}
    public StudentResponseMetadata(String keyVersion, int numCorrect, double percentCorrect, String id){
        this.keyVersion = keyVersion;
        this.numCorrect = numCorrect;
        this.percentCorrect = percentCorrect;
        this.id = id;
    }

    public String getKeyVersion() {
        return keyVersion;
    }

    public void setKeyVersion(String keyVersion) {
        this.keyVersion = keyVersion;
    }

    public int getNumCorrect() {
        return numCorrect;
    }

    public void setNumCorrect(int numCorrect) {
        this.numCorrect = numCorrect;
    }

    public double getPercentCorrect() {
        return percentCorrect;
    }

    public void setPercentCorrect(double percentCorrect) {
        this.percentCorrect = percentCorrect;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }
}
