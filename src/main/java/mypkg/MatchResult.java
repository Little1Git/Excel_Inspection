package mypkg;

import java.util.ArrayList;

public class MatchResult {
    public String key;
    public String result;
    public String actualValue;
    public ArrayList<String> expectedValue;
    public String expectedValueString;

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
                ", expectedValue=" + expectedValue +
                '}';
    }
}
