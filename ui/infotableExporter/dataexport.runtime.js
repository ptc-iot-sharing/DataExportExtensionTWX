(function() {
  var addedDefaultStyles = false;

  TW.Runtime.Widgets.infotableExporter = function() {

    var thisWidget = this;
    var roundedCorners = true;

    this.runtimeProperties = function() {
      return {
        'needsDataLoadingAndError': false,
        'propertyAttributes': {
          'Label': {
            'isLocalizable': true
          }
        }
      };
    };

    this.renderHtml = function() {

      var formatResult = TW.getStyleFromStyleDefinition(thisWidget.getProperty('Style', 'DefaultButtonStyle'));
      var formatResult2 = TW.getStyleFromStyleDefinition(thisWidget.getProperty('HoverStyle', 'DefaultButtonHoverStyle'));
      var formatResult3 = TW.getStyleFromStyleDefinition(thisWidget.getProperty('ActiveStyle', 'DefaultButtonActiveStyle'));
      var textSizeClass = 'textsize-normal';
      if (this.getProperty('Style') !== undefined) {
        textSizeClass = TW.getTextSizeClassName(formatResult.textSize);
      }

      var html =
        '<div class="dropdown widget-content widget-infotableExporter ">' +
        '<button class="dataexport-element2" tabindex="' +
        thisWidget.getProperty('TabSequence') + '">' +
        '<span class="widget-infotableExporter-icon">' +
        ((formatResult.image !== undefined && formatResult.image.length > 0) ?
          '<img class="default" src="' + formatResult.image + '"/>' : '') +
        ((formatResult2.image !== undefined && formatResult2.image.length > 0) ?
          '<img class="hover" src="' + formatResult2.image + '"/>' : '') +
        ((formatResult3.image !== undefined && formatResult3.image.length > 0) ?
          '<img class="active" src="' + formatResult3.image + '"/>' : '') +
        '</span>' + '<span class="widget-infotableExporter-text ' +
        textSizeClass + '">' +
        (thisWidget.getProperty('Label') === undefined ?
          'Export' : Encoder.htmlEncode(thisWidget.getProperty('Label'))) +
        '</span>' + '</button>' +
        '<div class="dropdown-content">' +
        '  <ul>\
      <li><a class="exportButton csvExport">CSV</a></li>\
      <li><a class="exportButton excelExport">Excel</a></li>\
      <li><a class="exportButton pdfExport">PDF</a></li>\
      <li><a class="exportButton wordExport">Word</a></li>\
      </ul>' +
        '</div></div> ';
      return html;
    };

    this.afterRender = function() {
      var formatResult = TW.getStyleFromStyleDefinition(thisWidget.getProperty('Style', 'DefaultButtonStyle'));
      var buttonHoverStyle = TW.getStyleFromStyleDefinition(thisWidget.getProperty('HoverStyle', 'DefaultButtonHoverStyle'));
      var buttonActiveStyle = TW.getStyleFromStyleDefinition(thisWidget.getProperty('ActiveStyle', 'DefaultButtonActiveStyle'));
      var buttonFocusStyle = TW.getStyleFromStyleDefinition(thisWidget.getProperty('FocusStyle', 'DefaultButtonFocusStyle'));

      var cssInfo = TW.getStyleCssTextualNoBackgroundFromStyle(formatResult);
      var cssButtonBackground = TW.getStyleCssGradientFromStyle(formatResult);
      var cssButtonBorder = TW.getStyleCssBorderFromStyle(formatResult);

      var cssButtonHoverText = TW.getStyleCssTextualNoBackgroundFromStyle(buttonHoverStyle);
      var cssButtonHoverBackground = TW.getStyleCssGradientFromStyle(buttonHoverStyle);
      var cssButtonHoverBorder = TW.getStyleCssBorderFromStyle(buttonHoverStyle);

      var cssButtonActiveText = TW.getStyleCssTextualNoBackgroundFromStyle(buttonActiveStyle);
      var cssButtonActiveBackground = TW.getStyleCssGradientFromStyle(buttonActiveStyle);
      var cssButtonActiveBorder = TW.getStyleCssBorderFromStyle(buttonActiveStyle);

      var cssButtonFocusBorder = TW.getStyleCssBorderFromStyle(buttonFocusStyle);

      roundedCorners = this.getProperty('RoundedCorners');
      if (roundedCorners === undefined) {
        roundedCorners = true;
      }

      if (roundedCorners == true) {
        thisWidget.jqElement.addClass('roundedCorners');
      }

      if (buttonHoverStyle.image.length === 0) {
        thisWidget.jqElement.addClass('singleImageOnly');
      }

      if (thisWidget.getProperty('Style', 'DefaultButtonStyle') === 'DefaultButtonStyle' && thisWidget.getProperty('HoverStyle', 'DefaultButtonHoverStyle') === 'DefaultButtonHoverStyle' && thisWidget.getProperty('ActiveStyle', 'DefaultButtonActiveStyle') === 'DefaultButtonActiveStyle' && thisWidget.getProperty('FocusStyle', 'DefaultButtonFocusStyle') === 'DefaultButtonFocusStyle') {
        if (!addedDefaultStyles) {
          addedDefaultStyles = true;
          var defaultStyles = '.widget-infotableExporter .dataexport-element2 { ' + cssButtonBackground + cssButtonBorder + ' }' +
            ' .widget-infotableExporter .dataexport-element2 span { ' + cssInfo + ' } ' +
            ' .widget-infotableExporter .dataexport-element2:hover { ' + cssButtonHoverBackground + cssButtonHoverBorder + ' }' +
            ' .widget-infotableExporter .dataexport-element2:hover span { ' + cssButtonHoverText + ' } ' +
            ' .widget-infotableExporter .dataexport-element2:active { ' + cssButtonActiveBackground + cssButtonActiveBorder + ' }' +
            ' .widget-infotableExporter .dataexport-element2:active span { ' + cssButtonActiveText + ' } ' +
            ' .widget-infotableExporter.focus .dataexport-element2 { ' + cssButtonFocusBorder + ' }';
          $.rule(defaultStyles).appendTo(TW.Runtime.globalWidgetStyleEl);
        }
      } else {
        var styleBlock =
          '<style>' +
          '#' + thisWidget.jqElementId + ' .dataexport-element2 { ' + cssButtonBackground + cssButtonBorder + ' } ' +
          '#' + thisWidget.jqElementId + ' .dataexport-element2:hover { ' + cssButtonHoverBackground + cssButtonHoverBorder + ' } ' +
          '#' + thisWidget.jqElementId + ' .dataexport-element2:active { ' + cssButtonActiveBackground + cssButtonActiveBorder + ' }' +
          '#' + thisWidget.jqElementId + ' .dataexport-element2 span { ' + cssInfo + ' } ' +
          '#' + thisWidget.jqElementId + ' .dataexport-element2:hover span { ' + cssButtonHoverText + ' } ' +
          '#' + thisWidget.jqElementId + ' .dataexport-element2:active span { ' + cssButtonActiveText + ' } ' +
          '#' + thisWidget.jqElementId + '.focus .dataexport-element2 { ' + cssButtonFocusBorder + ' }' +
          '</style>';

        $(styleBlock).prependTo(thisWidget.jqElement);
      }

      // not sure what this is all about, but I'm making it consistent with the button control ROE 9/28/12
      var buttonTextLineHeight = this.getProperty('Height');
      var buttonAdjustedTextLineHeight = buttonTextLineHeight - (formatResult.lineThickness * 2) + 'px';
      thisWidget.jqElement.find('.widget-infotableExporter-text').css('line-height', buttonAdjustedTextLineHeight);

      var dataBinding = this.getProperty('DataBinding');
      var triggerName = '';
      if (dataBinding === undefined) {
        TW.log.error('Data Export will not work since DataBinding is not set');
      } else {
        var section = dataBinding.SourceSection;
        var id = dataBinding.SourceId;
        if (section === undefined) {
          section = dataBinding.Section;
        }
        if (id === undefined) {
          id = dataBinding.Id;
        }
        triggerName = 'Data' + '_' + this.mashup.Data[section].DataName + '_' + id;
        // "Data_Logs_ConfigurationLog_GetLogEntries"
      }
      var idOfThisMashup = this.idOfThisMashup;
      var exportButton = thisWidget.jqElement.find(".exportButton");
      var exporterResource = "InfotableExporterFunctions";
      var invoker = new ThingworxInvoker({
        "entityType": "Resources",
        "entityName": "InfotableExporterFunctions",
        "characteristic": "Services",
        "target": "",
        "apiMethod": "post",
        "parameters": {}
      });
      var saveDataCallback = function(data) {
        window.location = data.result.rows[0].result;;
      };

      exportButton.bind('click', function(e) {
        if (triggerName.length > 0) {
          if (this.classList.contains("csvExport")) {
            TW.Runtime.exportingToCsv = true;
            $(idOfThisMashup).triggerHandler(triggerName);
            TW.Runtime.exportingToCsv = false;
          } else {
            invoker.setParameterValue("infotable", thisWidget.getProperty('Data'));
            if (this.classList.contains("excelExport")) {
              invoker.setParameterValue("target", "ExportInfotableAsExcel")
            } else if (this.classList.contains("pdfExport")) {
              invoker.setParameterValue("target", "ExportInfotableAsPdf")
            } else if (this.classList.contains("wordExport")) {
              invoker.setParameterValue("target", "ExportInfotableAsWord")
            }
             invoker.invokeService(saveDataCallback, function() {
                        alert("Failed to get export url. Please try again.")
                      });
          }

        }
        e.preventDefault();
      });

      var iconAlignment = this.getProperty('IconAlignment');
      var iconElement = thisWidget.jqElement.find('.widget-infotableExporter-icon');
      var buttonText = thisWidget.jqElement.find('.widget-infotableExporter-text');

      if (iconAlignment == 'right') {
        $(iconElement).insertAfter(buttonText);
        thisWidget.jqElement.addClass('iconRight');
      }

      var widgetSelector = '#' + this.jqElementId + ' .dataexport-element2';
      var widgetContainer = '#' + this.jqElementId;

      $(widgetSelector).on('focusin', function() {
        $(widgetContainer).addClass('focus');
      });

      $(widgetSelector).on('blur', function(e) {
        $(widgetContainer).removeClass('focus');
      });

    };

    this.updateProperty = function(updatePropertyInfo) {
      if (updatePropertyInfo.TargetProperty === "Data") {
        thisWidget.setProperty("Data", updatePropertyInfo.RawDataFromInvoke);
      }
    };

    this.beforeDestroy = function() {
      try {
        thisWidget.jqElement.unbind();
      } catch (err) {
        TW.log.error('Error in TW.Runtime.Widgets.dataexport.beforeDestroy', err);
      }
    };

  };
}());
