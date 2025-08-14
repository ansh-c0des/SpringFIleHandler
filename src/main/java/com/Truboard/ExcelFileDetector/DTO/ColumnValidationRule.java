
package com.Truboard.ExcelFileDetector.DTO;

import lombok.Data;

@Data
public class ColumnValidationRule {
    private String type; // "date", "number", "text"
    private String format; // date format if type=date
    private boolean required;
    private Double min;
    private Double max;
    private String regex;

    // getters & setters
    public String getType() { return type; }
    public void setType(String type) { this.type = type; }
    public String getFormat() { return format; }
    public void setFormat(String format) { this.format = format; }
    public boolean isRequired() { return required; }
    public void setRequired(boolean required) { this.required = required; }
    public Double getMin() { return min; }
    public void setMin(Double min) { this.min = min; }
    public Double getMax() { return max; }
    public void setMax(Double max) { this.max = max; }
    public String getRegex() { return regex; }
    public void setRegex(String regex) { this.regex = regex; }
}
