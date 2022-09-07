package setting;

import java.util.ArrayList;

public class Setting {

    private String setting;
    private String name;
    private String value;
    private String startRange;
    private String finishRange;
    private String step;
    private boolean right = true;
    private boolean addASU = false;
    private ArrayList<String> allValues = new ArrayList<>();
    private String type = "";


    public String getSetting() {
        return setting;
    }

    public void setSetting(String setting) {
        this.setting = setting;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value.replace(".", ",");
    }

    public String getStartRange() {
        return startRange;
    }

    public void setStartRange(String startRange) {
        this.startRange = startRange.replace(".", ",");
    }

    public String getFinishRange() {
        return finishRange;
    }

    public void setFinishRange(String finishRange) {
        this.finishRange = finishRange.replace(".", ",");
    }

    public String getStep() {
        return step;
    }

    public void setStep(String step) {
        String[] intStep = step.split("\\.");
        if (intStep.length == 2) {
            int countZero = intStep[1].length();
            if (value.split(",").length == 2) {
                int differentZero = countZero - value.split(",")[1].length();
                for (int i = 0; i < differentZero; i++) {
                    value += "0";
                }
            } else if (value.split(",").length == 1) {
                value += ",";
                for (int i = 0; i < countZero; i++) {
                    value += "0";
                }
            }


            if (startRange.split(",").length == 2) {
                int differentZero = countZero - startRange.split(",")[1].length();
                for (int i = 0; i < differentZero; i++) {
                    startRange += "0";
                }
            } else if (startRange.split(",").length == 1) {
                startRange += ",";
                for (int i = 0; i < countZero; i++) {
                    startRange += "0";
                }
            }


            if (finishRange.split(",").length == 2) {
                int differentZero = countZero - finishRange.split(",")[1].length();
                for (int i = 0; i < differentZero; i++) {
                    finishRange += "0";
                }
            } else if (finishRange.split(",").length == 1) {
                finishRange += ",";
                for (int i = 0; i < countZero; i++) {
                    finishRange += "0";
                }
            }
        }
        this.step = step.replace(".", ",");
    }


    public boolean isAddASU() {
        return addASU;
    }

    public void setAddASU(boolean addASU) {
        this.addASU = addASU;
    }


    public String getRange() {
        return startRange + "-" + finishRange;
    }

    public boolean isRight() {
        return right;
    }

    public void setRight(boolean right) {
        this.right = right;
    }

    public ArrayList<String> getAllValues() {
        return allValues;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    @Override
    public String toString() {
        return "setting.Setting{" +
                "setting='" + setting + '\'' +
                ", name='" + name + '\'' +
                ", value=" + value +
                ", startRange=" + startRange +
                ", finishRange=" + finishRange +
                ", step=" + step +
                '}';
    }
}
