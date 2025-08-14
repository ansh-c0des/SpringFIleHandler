package com.Truboard.ExcelFileDetector.config;

import com.Truboard.ExcelFileDetector.DTO.ColumnValidationRule;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Configuration
@ConfigurationProperties(prefix = "excel")
public class ExcelValidationConfig {
    private Map<String, ColumnValidationRule> validations;
    private List<String> requiredColumns = new ArrayList<>();

    public Map<String, ColumnValidationRule> getValidations() {
        return validations;
    }

    public void setValidations(Map<String, ColumnValidationRule> validations) {
        this.validations = validations;
    }

    public List<String> getRequiredColumns() {
        return requiredColumns;
    }

    public void setRequiredColumns(List<String> requiredColumns) {
        this.requiredColumns = requiredColumns;
    }
}
