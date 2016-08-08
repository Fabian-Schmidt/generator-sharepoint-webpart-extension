"use strict";
import * as $ from "jquery";
/**
 * Based on AppPartPropertyUIOverride from OfficeDev PnP.
 * https://github.com/OfficeDev/PnP/blob/master/Samples/Core.AppPartPropertyUIOverride/Core.AppPartPropertyUIOverrideWeb/Scripts/Contoso.AppPartPropertyUIOverride.js
 */
export module AppPartPropertyUIOverride {
    export interface contentSettings {
        category: string;
        optionalName?: string;
        optionalToolTip?: string;
        outputSeparator: boolean;
    }
    export class AppPartPropertyUIOverride {
        private isActive: boolean = undefined;
        private isActiveDeffered: JQueryDeferred<boolean> = $.Deferred();
		/**
		 * Create a new instance.
                 * @param webPartId Guid of the web part in the format 'cf3a0b82-b883-43ad-8ca8-affacfad8e83'.
		 */
        constructor(webPartId: string) {
            var that = this;
            if (!webPartId || webPartId.length == 0)
                throw 'Invalid webPartId';

            if (document.readyState == "complete"
                || document.readyState == "loaded"
                || document.readyState == "interactive") {
                // document has at least been parsed
                that.init(webPartId);
            } else {
                document.addEventListener("DOMContentLoaded", function(event) {
                    if (document.readyState == "complete"
                        || document.readyState == "loaded"
                        || document.readyState == "interactive") {
                        that.init(webPartId);
                    }
                });
            }
        }

        private init(webPartId: string): void {
            if (this.isActive === undefined) {
                var appPartPropertyPaneTd = document.getElementById("MSOTlPn_Parts");
                //var webPartTollPartId = 'ToolPartctl00_MSOTlPn_EditorZone_Edit1g_' + webPartId.toLowerCase().replace(/-/g, '_');
                //var webPartTollPart = document.getElementById(webPartTollPartId);
                this.isActive = false;
                if (appPartPropertyPaneTd != null) {
                    this.appPartPropertyPaneTdjQueryWrapper = $(appPartPropertyPaneTd);

                    var containsWebPartTollPart = appPartPropertyPaneTd.innerHTML.indexOf(webPartId.toLowerCase().replace(/-/g, '_'));
                    //This class is active when edit tool pane for this app part is visible.
                    this.isActive = appPartPropertyPaneTd != null && containsWebPartTollPart > -1;
                }
                this.isActiveDeffered.resolve(this.isActive);
            }
        }

        /**
         * Check if property ui override is active.
         * Edit tool pane was found for this app part.
         */
        public IsActive(): JQueryPromise<boolean> { return this.isActiveDeffered.promise(); }

		/**
		 * Creates a new html content area in the App Part Property UI category bottom and returns a jQuery object that wraps the new content area created.
                 * @param settings A JavaScript object that contains the required and optional settings for this operation. {category: "The Category"} is required.Optional object properties are: optionalName, optionalToolTip, and outputSeparator 
		 * @returns A jQuery object that wraps the new content area created.
                 */
        public createNewContentAtBottom(settings: contentSettings): JQuery {
            var actualSettings = {
                category: "",
                optionalName: "",
                optionalToolTip: "",
                outputSeparator: true
            };

            // merge user supplied settings with default settings
            $.extend(actualSettings, settings);

            // get the category content table jQuery wrapper
            var categoryContentTable = this.getCategoryContentTable(actualSettings.category);

            // increment the new property counter
            this.newPropertyCounter = this.newPropertyCounter + 1;
            var newPropertyCounter = this.newPropertyCounter;

            // get last td
            categoryContentTable.find("td:last").append("<div class=\"UserDottedLine\" style=\"width: 100%;\"></div>");

            // build html string to inject
            var html: string[] = [];
            html.push("<tr><td>");

            if (actualSettings.optionalName !== null && actualSettings.optionalName !== "") {
                html.push("<div class=\"UserSectionHead\"><label title=\"" + actualSettings.optionalToolTip + "\">" + actualSettings.optionalName + "</label></div>");
            }

            html.push("<div class=\"UserSectionBody\" id=\"AppPartPropertyUINewContentArea" + newPropertyCounter + "\" style=\"margin-bottom: 10px;\"></div>");

            html.push("</td></tr>");

            categoryContentTable.append(html.join(""));
            return categoryContentTable.find("#AppPartPropertyUINewContentArea" + newPropertyCounter);
        }

        /**
         * Public static function that creates a new html content area in the specified App Part Property UI category top and returns a jQuery object that wraps the new content area created.
         * @param settings A JavaScript object that contains the required and optional settings for this operation.  {category: "The Category"} is required.Optional object properties are: optionalName, optionalToolTip, and outputSeparator
         * @returns A jQuery object that wraps the new content area created.
         */
        public createNewContentAtTop(settings: AppPartPropertyUIOverride.contentSettings): JQuery {
            var actualSettings = {
                category: "",
                optionalName: "",
                optionalToolTip: "",
                outputSeparator: true
            };

            // merge user supplied settings with default settings
            $.extend(actualSettings, settings);

            // get the category content table jQuery wrapper
            var categoryContentTable = this.getCategoryContentTable(actualSettings.category);

            // increment the new property counter
            this.newPropertyCounter = this.newPropertyCounter + 1;
            var newPropertyCounter = this.newPropertyCounter;

            // build html string to inject
            var html: string[] = [];
            html.push("<tr><td>");

            if (actualSettings.optionalName !== null && actualSettings.optionalName !== "") {
                html.push("<div class=\"UserSectionHead\"><label title=\"" + actualSettings.optionalToolTip + "\">" + actualSettings.optionalName + "</label></div>");
            }

            html.push("<div class=\"UserSectionBody\" id=\"AppPartPropertyUINewContentArea" + newPropertyCounter + "\" style=\"margin-bottom: 10px;\"></div><div class=\"UserDottedLine\" style=\"width: 100%;\"></div>");

            html.push("</td></tr>");

            categoryContentTable.prepend(html.join(""));
            return categoryContentTable.find("#AppPartPropertyUINewContentArea" + newPropertyCounter);
        }

        /**
         * Public static function that expands the specified category and closes all others in the App Part property pane UI.
         * @param category The category display text of the category to expand.
         * Example: "Custom Category 1"
         */
        public expandCategory(category: string) {
            function ensureCategoryOpened(opened: boolean, categoryJQueryWrapper: JQuery): void {
                // first determined the current state (is opened?)
                var currentStateIsOpened = false;
                var atag = categoryJQueryWrapper.find(".UserSectionTitle a:first");
                if (atag.html().indexOf("/TPMax2.gif") === -1) {
                    currentStateIsOpened = true;
                }

                // we got the current state... now compare it to what we want it to be
                if (currentStateIsOpened !== opened) {
                    // the current state is different than what we want it to be
                    // trigger a click on the UI
                    atag.trigger("click");
                }
            }

            // ensure the custom category dictionary is present
            this.ensureCategoryDictionaryPresent();

            // now loop through all categories and ensure the correct ones are closed/opened
            var categories = this.categoryDictionary;
            var categoryKeys = Object.keys(categories);
            var categoryKey: string = null;

            for (var i = categoryKeys.length - 1; i > -1; i = i - 1) {
                categoryKey = categoryKeys[i];
                if (categoryKey === category) {
                    ensureCategoryOpened(true, categories[categoryKey]);
                } else {
                    ensureCategoryOpened(false, categories[categoryKey]);
                }
            }
        }

        // /**
        //  * Tells the App Part property UI framework you are done with overriding the App Part property UI and to show the property pane again.
        //  */
        // public finished() {
        //         this.appPartPropertyPaneTdjQueryWrapper.show();
        // }

        /**
         * Public static function that gets the current value of the declared string, number, enum, or boolean property that's already rendered in the App Part property UI.
         * @param name The display name of the property.  Example: "My Setting 1"
         * @param category The display name of the category.  Example: "Custom Category 1"
         * @returns The value of that property if it exists.
         */
        public getValue(name: string, category: string): any {
            var inputElementjQueryWrapper = this.getInputElementJQueryWrapper(name, category);
            var dataType = this.getInputElementDataType(name, category);
            var returnValue: any = null;
            switch (dataType) {
                case "CHECKBOX":
                    returnValue = inputElementjQueryWrapper.is(":checked");
                    break;
                case "SELECT":
                    returnValue = inputElementjQueryWrapper.val();
                    break;
                case "TEXT":
                    returnValue = inputElementjQueryWrapper.val();
                    if (!isNaN(returnValue)) {
                        returnValue = returnValue * 1;
                    }
                    break;
            }

            return returnValue;
        }

        /**
         * Public static function that hides the specified property in the App Part property UI.
         * @param name The display name of the property.  Example: "My Property 1"
         * @param category The display name of the category.  Example: "Custom Category 1"
         */
        public hideProperty(name: string, category: string) {
            this.ensureCategoryDictionaryPresent();
            var labelElement: HTMLLabelElement = null;
            var bodyDivjQueryWrapper: JQuery = null;
            var labelArray: JQuery = this.categoryDictionary[category].find("label");
            var parentElement: JQuery = null;
            for (var i = 0; i < labelArray.length; i = i + 1) {
                labelElement = <HTMLLabelElement>labelArray[i];
                if (labelElement.innerHTML === name || labelElement.innerText === name || labelElement.htmlFor.indexOf(name) > 50) {
                    parentElement = $(labelElement).parent().parent();

                    // found parent
                    // now hide it
                    parentElement.hide();
                    break;
                }
            }
        }

        /**
         * Public static function that moves the specified category to the top of the App Part property pane UI.
         * @param category The display name of the category.  Example: "Custom Category 1"
         */
        public moveCategoryToTop(category: string) {
            // find the category and parent
            this.ensureCategoryDictionaryPresent();
            var sourceDivToMove = this.categoryDictionary[category];
            var parentDivToMoveUnder = this.parentCategoryDivJQueryWrapper;

            // do the move
            sourceDivToMove.prependTo(parentDivToMoveUnder);

            // reload the dictionary
            this.reloadCategoryDictionary();
        }

        /**
         * Public static function that sets the current value of the declared string, number, enum, or boolean property that's already rendered in the App Part property UI.
         * @param name The display name of the property.  Example: "My Setting 1"
         * @param value The value to set.  Example: "The Value"
         * @param category The display name of the category.  Example: "Custom Category 1"
         */
        public setValue(name: string, value: any, category: string) {
            var inputElementjQueryWrapper = this.getInputElementJQueryWrapper(name, category);
            var dataType = this.getInputElementDataType(name, category);
            switch (dataType) {
                case "CHECKBOX":
                    inputElementjQueryWrapper.attr("checked", value);
                    break;
                case "SELECT":
                    inputElementjQueryWrapper.val(value);
                    break;
                case "TEXT":
                    inputElementjQueryWrapper.val(value + "");
                    break;
            }
        }

        /**
         * Public static function that renders tool tips as html instruction text below each property in the specified category.
         * @param category The display name of the category.  Example: "Custom Category 1"
         */
        public renderToolTipsAsInstructions(category: string) {
            this.ensureCategoryDictionaryPresent();
            var labelElement: JQuery = null;
            var bodyDivjQueryWrapper: JQuery = null;
            var labelArray: JQuery = this.categoryDictionary[category].find("label");
            var parentElement: JQuery = null;
            var toolTip = "";
            var labeljQueryWrapper: JQuery = null;
            var userDottedLineDiv: JQuery = null;
            for (var i = 0; i < labelArray.length; i = i + 1) {
                labelElement = $(labelArray[i]);
                if (labelElement.attr("style") !== "display: none;") {
                    parentElement = labelElement.parent().parent();
                    if (parentElement[0].tagName.toUpperCase() === "TD") {
                        // now we have a property TD
                        // see if we have a tool tip
                        toolTip = "";
                        labeljQueryWrapper = parentElement.find("div.UserSectionHead").find("label:first");
                        if (labeljQueryWrapper.length > 0) {
                            toolTip = labeljQueryWrapper.attr("title");
                            if (toolTip.length > 0) {
                                // see if there's UserDottedLine div
                                userDottedLineDiv = parentElement.find("div.UserDottedLine");
                                if (userDottedLineDiv.length > 0) {
                                    // need to add it before this node
                                    $(userDottedLineDiv[0]).before("<div style=\"margin-bottom: 10px;color: #cccccc;font-size: smaller;font-style:italic\" class=\"appPartInstruction\">" + toolTip + "</div>");
                                } else {
                                    // need to add it to end of parent content
                                    parentElement.append("<div style=\"margin-bottom: 10px;color: #cccccc;font-size: smaller;font-style:italic\" class=\"appPartInstruction\">" + toolTip + "</div>");
                                }
                            }
                        }
                    }
                }
            }
        }

        private categoryDictionary: { [category: string]: JQuery } = null;
        private parentCategoryDivJQueryWrapper: JQuery = null;
        private appPartPropertyPaneTdjQueryWrapper: JQuery = null;
        private newPropertyCounter: number = 0;

        private ensureCategoryDictionaryPresent(): void {
            if (this.categoryDictionary === null) {
                this.reloadCategoryDictionary();
            }
        }
        private getCategoryContentTable(category: string): JQuery {
            this.ensureCategoryDictionaryPresent();
            return $(this.categoryDictionary[category].find("div.ms-propGridTable > table")[0]);
        }

        private reloadCategoryDictionary(): void {
            var categoryDictionary: { [category: string]: JQuery } = {};
            var atags: JQuery = this.appPartPropertyPaneTdjQueryWrapper.find("div.UserSectionTitle a");
            var atag: JQuery = null;
            var category: string = null;
            var categoryDivJQueryWrapper: JQuery = null;
            var parentCategoryDiv: JQuery = null;

            for (var i = 0; i < atags.length; i = i + 1) {
                atag = $(atags[i]);
                category = atag.html();
                if (category.indexOf("<img") === -1) {
                    category = $.trim(category.replace("&nbsp;", " "));
                    categoryDivJQueryWrapper = atag.parentsUntil("div.ms-TPBody");

                    if (parentCategoryDiv === null) {
                        parentCategoryDiv = categoryDivJQueryWrapper.last().parent().parent();
                    }

                    if (!categoryDictionary.hasOwnProperty(category)) {
                        categoryDictionary[category] = categoryDivJQueryWrapper.last();
                    }
                }
            }

            this.parentCategoryDivJQueryWrapper = parentCategoryDiv;
            this.categoryDictionary = categoryDictionary;
        }

        private inputElementDataTypeDictionary: { [categoryAndName: string]: string } = {};
        private inputElementJQueryDictionary: { [categoryAndName: string]: JQuery } = {};

        private getInputElementDataType(name: string, category: string): string {
            var key = category + "-" + name;
            if (this.inputElementDataTypeDictionary.hasOwnProperty(key)) {
                return this.inputElementDataTypeDictionary[key];
            } else {
                this.getInputElementJQueryWrapper(name, category);
                return this.inputElementDataTypeDictionary[key];
            }
        }
        private getInputElementJQueryWrapper(name: string, category: string): JQuery {
            var key = category + "-" + name;
            if (this.inputElementJQueryDictionary.hasOwnProperty(key)) {
                return this.inputElementJQueryDictionary[key];
            } else {
                this.ensureCategoryDictionaryPresent();
                var labelElement: HTMLLabelElement = null;
                var bodyDivjQueryWrapper: JQuery = null;
                var labelArray: JQuery = this.categoryDictionary[category].find("label");
                var parentElement: JQuery = null;
                for (var i = 0; i < labelArray.length; i = i + 1) {
                    labelElement = <HTMLLabelElement><any>labelArray[i];
                    if (labelElement.innerHTML === name || labelElement.innerText === name || labelElement.htmlFor.indexOf(name) > 50) {
                        parentElement = $(labelElement).parent().parent();
                        bodyDivjQueryWrapper = parentElement.find("div.UserSectionBody");

                        if (bodyDivjQueryWrapper.length === 0) {
                            returnValue = parentElement.find("div.UserSectionHead > span > input:checkbox");
                            if (returnValue.length > 0) {
                                dataType = "CHECKBOX";
                                returnValue = $(returnValue[0]);
                                this.inputElementJQueryDictionary[key] = returnValue;
                                this.inputElementDataTypeDictionary[key] = dataType;
                                return returnValue;
                            } else {
                                return null;
                            }
                        } else {
                            var inputElements = bodyDivjQueryWrapper.find(":input");
                            var inputElement: HTMLInputElement = null;
                            var returnValue: JQuery = null;
                            var dataType: string = null;
                            if (inputElements.length > 0) {
                                for (var j = 0; j < inputElements.length; j = j + 1) {
                                    inputElement = <HTMLInputElement><any>inputElements[j];
                                    dataType = inputElement.tagName.toUpperCase();
                                    switch (dataType) {
                                        case "INPUT":
                                            dataType = inputElement.getAttribute("type").toUpperCase();
                                            if (dataType !== "HIDDEN") {
                                                returnValue = $(inputElement);
                                                break;
                                            }
                                        case "SELECT":
                                            returnValue = $(inputElement);
                                            break;
                                    }

                                    if (returnValue) {
                                        break;
                                    }
                                }
                            }

                            if (returnValue) {
                                this.inputElementJQueryDictionary[key] = returnValue;
                                this.inputElementDataTypeDictionary[key] = dataType;
                                return returnValue;
                            } else {
                                return null;
                            }
                        }
                    }
                }
            }
        }

        private ensureNotNullString(value: string): string {
            if (typeof value === "string") {
                return value;
            } else {
                return "";
            }
        }
    }
}