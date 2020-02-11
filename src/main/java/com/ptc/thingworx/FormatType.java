package com.ptc.thingworx;

import com.thingworx.types.BaseTypes;

public enum FormatType {
    TEXT,
    INTEGER,
    FLOAT,
    DATE,
    MONEY,
    PERCENTAGE,
    HEADER;

    public static FormatType getFormatType(BaseTypes baseType) {
        if (baseType == BaseTypes.INTEGER || baseType == BaseTypes.LONG) {
            return FormatType.INTEGER;
        } else if (baseType == BaseTypes.NUMBER) {
            return FormatType.FLOAT;
        } else if (baseType == BaseTypes.DATETIME) {
            return FormatType.DATE;
        } else {
            return FormatType.TEXT;
        }
    }
}