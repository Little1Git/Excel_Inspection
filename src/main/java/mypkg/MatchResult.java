package mypkg;

import java.util.ArrayList;

public class MatchResult {
    public String key;
    public boolean result;
    public String actualValue;
    public ArrayList<String> expectedValue;
    public String expectedValueInString;

    // 构造函数
    public MatchResult() {
        this.expectedValue = new ArrayList<>();
    }

    @Override
    public String toString() {
        return "MatchResult{" +
                "key='" + key + '\'' +
                ", result='" + result + '\'' +
                ", actualValue='" + actualValue + '\'' +
                ", expectedValueString='" + expectedValueInString + '\'' +
                ", expectedValue=" + expectedValue +
                '}';
    }
}
