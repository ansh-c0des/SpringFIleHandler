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
    
    // Error highlighting configuration
    private ErrorHighlightConfig errorHighlight = new ErrorHighlightConfig();

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

    public ErrorHighlightConfig getErrorHighlight() {
        return errorHighlight;
    }

    public void setErrorHighlight(ErrorHighlightConfig errorHighlight) {
        this.errorHighlight = errorHighlight;
    }

    public static class ErrorHighlightConfig {
        private String backgroundColor = "RED";
        private String commentAuthor = "Excel Validator";
        private String commentPrefix = "Validation Error:\n";
        private boolean showCurrentValue = true;
        private int commentWidth = 3; // columns
        private int commentHeight = 3; // rows

        public String getBackgroundColor() {
            return backgroundColor;
        }

        public void setBackgroundColor(String backgroundColor) {
            this.backgroundColor = backgroundColor;
        }

        public String getCommentAuthor() {
            return commentAuthor;
        }

        public void setCommentAuthor(String commentAuthor) {
            this.commentAuthor = commentAuthor;
        }

        public String getCommentPrefix() {
            return commentPrefix;
        }

        public void setCommentPrefix(String commentPrefix) {
            this.commentPrefix = commentPrefix;
        }

        public boolean isShowCurrentValue() {
            return showCurrentValue;
        }

        public void setShowCurrentValue(boolean showCurrentValue) {
            this.showCurrentValue = showCurrentValue;
        }

        public int getCommentWidth() {
            return commentWidth;
        }

        public void setCommentWidth(int commentWidth) {
            this.commentWidth = commentWidth;
        }

        public int getCommentHeight() {
            return commentHeight;
        }

        public void setCommentHeight(int commentHeight) {
            this.commentHeight = commentHeight;
        }
    }
}