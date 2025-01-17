package mypkg;

import java.util.ArrayList;

public class MatchResult {
    public String key;
    public String result;
    public String actualValue;
    public ArrayList<String> expectedValue;

    // 构造函数
    public MatchResult() {
        this.expectedValue = new ArrayList<>();
    }
}
