﻿TW.IDE.Widgets.infotableExporter = function () {
	var roundedCorners = true;
    this.widgetProperties = function () {
        return {
            'name': 'Infotable Exporter',
            'description': 'Enables you to any infotable to CSV/Excel/Word/PDF',
            'category': ['Data'],
            'defaultBindingTargetProperty': 'Data',
            //'customEditor': 'infotableExporterCustomEditor',
            //'customEditorMenuText': 'Configure Data Export Fields',
            'properties': {
                'Label': {
                    'description': 'The text that appears on the button',
                    'defaultValue': 'Export',
                    'baseType': 'STRING',
                    'isLocalizable': true
                },
                'ColumnFormat': {
                    'isVisible': false,
                    'baseType': 'STRING'
                },
                'DataBinding': {
                    'isVisible': false,
                    'baseType': 'JSON'
                },
                'TabSequence': {
                    'description': 'Tab sequence index',
                    'baseType': 'NUMBER',
                    'defaultValue': 0
                },
                'RoundedCorners': {
                    'description': 'Do you want the corners on the button rounded',
                    'baseType': 'BOOLEAN',
                    'defaultValue': true
                },
                'Width': {
                    'description': 'width of widget',
                    'baseType': 'NUMBER',
                    'defaultValue': 75
                },
                'Height': {
                    'description': 'height of widget',
                    'baseType': 'NUMBER',
                    'defaultValue': 30
                    // 'isEditable': false
                },
                'Data': {
                    'description': 'Select an infotable as the data source for this property',
                    'isBindingTarget': true,
                    'isEditable': false,
                    'baseType': 'INFOTABLE',
                    'warnIfNotBoundAsTarget': true
                },
				'Style': {
                    'baseType': 'STYLEDEFINITION',
                    'defaultValue': 'DefaultButtonStyle'
                },
				'HoverStyle': {
                    'baseType': 'STYLEDEFINITION',
					'defaultValue': 'DefaultButtonHoverStyle'
                },
				'ActiveStyle': {
                    'baseType': 'STYLEDEFINITION',
					'defaultValue': 'DefaultButtonActiveStyle'
                },
				'FocusStyle': {
                    'baseType': 'STYLEDEFINITION',
					'defaultValue': 'DefaultButtonFocusStyle'
                },
                'IconAlignment': {
                    'description': 'Either align the icon for the button to the left or the right of the text',
                    'baseType': 'STRING',
                    'defaultValue': 'left',
                    'selectOptions': [
                        { value: 'left', text: 'Left' },
                        { value: 'right', text: 'Right' }
                    ]
                }
            }
        };
    };

    this.renderHtml = function () {

		var formatResult = TW.getStyleFromStyleDefinition(this.getProperty('Style', 'DefaultButtonStyle'));
		var formatResult2 = TW.getStyleFromStyleDefinition(this.getProperty('HoverStyle', 'DefaultButtonHoverStyle'));
		var formatResult3 = TW.getStyleFromStyleDefinition(this.getProperty('ActiveStyle', 'DefaultButtonActiveStyle'));

	    var textSizeClass = 'textsize-normal';
	    if (this.getProperty('Style') !== undefined) {
	        textSizeClass = TW.getTextSizeClassName(formatResult.textSize);
	    }

		var html = '';

		var buttonBorderWidth = TW.getStyleCssBorderWidthOnlyFromStyle(formatResult);
		var buttonHeight = this.getProperty('Height');
		var adjustedWrapperHeight = this.getProperty('Height') - buttonBorderWidth * 2;

        html +=
            '<div class="widget-content widget-infotableExporter">'
				+ '<div class="widget-infotableExporter-wrapper">'
					+ '<div class="widget-infotableExporter-element" style="height:'+ buttonHeight +'px; line-height:'+ adjustedWrapperHeight +'px;">'
                        + '<span class="widget-infotableExporter-icon">'
    						+ ((formatResult.image !== undefined && formatResult.image.length > 0) ? '<img class="default" src="' + formatResult.image + '"/>' : '')
    						+ ((formatResult2.image !== undefined && formatResult2.image.length > 0) ? '<img class="hover" src="' + formatResult2.image + '"/>' : '')
    						+ ((formatResult3.image !== undefined && formatResult3.image.length > 0) ? '<img class="active" src="' + formatResult3.image + '"/>' : '')
                        + '</span>'
						+ '<span class="widget-infotableExporter-text ' + textSizeClass + '">' + (this.getProperty('Label') === undefined ? 'Export' : Encoder.htmlEncode(this.getProperty('Label'))) + '</span>'
					+ '</div>'
				+ '</div>'
          + '</div>';
        return html;
    };

	this.afterRender = function () {
		var thisWidget = this;

		var buttonStyle = TW.getStyleFromStyleDefinition(thisWidget.getProperty('Style'));
		var buttonHoverStyle = TW.getStyleFromStyleDefinition(thisWidget.getProperty('HoverStyle'));
		var buttonActiveStyle = TW.getStyleFromStyleDefinition(thisWidget.getProperty('ActiveStyle'));

		var buttonBackground = TW.getStyleCssGradientFromStyle(buttonStyle);
		var buttonText = TW.getStyleCssTextualNoBackgroundFromStyle(buttonStyle);
		var buttonBorder = TW.getStyleCssBorderFromStyle(buttonStyle);
		var buttonHoverBG = TW.getStyleCssGradientFromStyle(buttonHoverStyle);
		var buttonHoverText = TW.getStyleCssTextualNoBackgroundFromStyle(buttonHoverStyle);
		var buttonHoverBorder = TW.getStyleCssBorderFromStyle(buttonHoverStyle);
		var cssButtonActiveText = TW.getStyleCssTextualNoBackgroundFromStyle(buttonActiveStyle);
		var cssButtonActiveBackground = TW.getStyleCssGradientFromStyle(buttonActiveStyle);
		var cssButtonActiveBorder = TW.getStyleCssBorderFromStyle(buttonActiveStyle);

		roundedCorners = this.getProperty('RoundedCorners');
		if (roundedCorners === undefined) {
			roundedCorners = true;
		}

		if (roundedCorners == true) {
			thisWidget.jqElement.addClass('roundedCorners');
		}

		thisWidget.jqElement.mousedown(function() {
    		thisWidget.jqElement.addClass('active');
		}).mouseup(function(){
			thisWidget.jqElement.removeClass('active');
		})

		var resource = TW.IDE.getMashupResource();
        var widgetStyles =
            '#' + thisWidget.jqElementId + '.widget-infotableExporter span {'+ buttonText + '} ' +
            '#' + thisWidget.jqElementId + '.widget-infotableExporter:hover span {'+ buttonHoverText + '} ' +
            '#' + thisWidget.jqElementId + '.widget-infotableExporter .widget-infotableExporter-element {'+ buttonBackground + buttonBorder + '} ' +
            '#' + thisWidget.jqElementId + '.widget-infotableExporter:hover .widget-infotableExporter-element {'+ buttonHoverBG + buttonHoverBorder + '} ' +
            '#' + thisWidget.jqElementId + '.widget-infotableExporter.active .widget-infotableExporter-element {'+ cssButtonActiveBackground + cssButtonActiveBorder +'}';
        resource.styles.append(widgetStyles);

        var iconAlignment = this.getProperty('IconAlignment');
        var iconElement = thisWidget.jqElement.find('.widget-infotableExporter-icon');
        var buttonText = thisWidget.jqElement.find('.widget-infotableExporter-text');

        if (iconAlignment == 'right') {
            $(iconElement).insertAfter(buttonText);
            thisWidget.jqElement.addClass('iconRight');
        }

	};

    this.afterAddBindingSource = function (bindingInfo) {
        if (bindingInfo['targetProperty'] === 'Data') {

            var dataSourceInfo = undefined;
            // get the infoTableDataShape associated with this property
            var infoTableDataShape = this.getInfotableMetadataForProperty('Data', undefined, function (binding) { dataSourceInfo = binding });
            this.setProperty('ColumnFormat', JSON.stringify(infoTableDataShape));
            this.setProperty('DataBinding', dataSourceInfo);
        }
    }

    this.afterSetProperty = function (name, value) {
        var result = false;
        switch (name) {
        case 'Label' :
    	case 'Style':
    	case 'Width':
		case 'Height':
        case 'RoundedCorners':
        case 'HoverStyle':
        case 'IconAlignment':
            result = true;
            break;
        }
        return result;
    };

		this.widgetIconUrl = function () {
				return "../Common/thingworx/widgets/dataexport/dataexport.ide.png";
		};
};
