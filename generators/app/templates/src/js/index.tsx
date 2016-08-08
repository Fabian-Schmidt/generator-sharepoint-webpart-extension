import * as React from "react";
import * as ReactDOM from "react-dom";
import * as $ from "jquery";

import { Counter } from "./components/Counter";
import * as AppPartPropertyUIOverride from "./AppPartPropertyUIOverride";

export module COMPANY.<%= classname %> {

    export interface ISetting {
        /**
         * Settings are valid.
         */
        v: boolean;
        /**
         * Selected label value.
         */
        l?: string;
    }

    export class <%= classname %> {
        propertyUIOverride: AppPartPropertyUIOverride.AppPartPropertyUIOverride.AppPartPropertyUIOverride = null;
        output: HTMLDivElement;
        webPartId: string;
        settings: ISetting;

        constructor(webPartDOMObject: HTMLDivElement) {
            this.output = webPartDOMObject;
            if (webPartDOMObject.parentElement.attributes['webpartid2' as any])
                this.webPartId = webPartDOMObject.parentElement.attributes['webpartid2' as any].value;
            else
                this.webPartId = webPartDOMObject.parentElement.attributes['webpartid' as any].value;
            try {
                this.settings = JSON.parse(window.atob(webPartDOMObject.parentElement.attributes['helplink'as any].value));
            } catch (err) {
                //init default settings
                this.settings = { v: false, l: 'Typescript' };
            }
            this.propertyUIOverride = new AppPartPropertyUIOverride.AppPartPropertyUIOverride.AppPartPropertyUIOverride(this.webPartId);
            this.propertyUIOverride.IsActive().done(this.PropertyUIOverrideActive.bind(this));

            this.Init();
        }

        PropertyUIOverrideActive(isActive: boolean) {
            var that = this;
            if (isActive) {
                var contentSettings = { category: 'Advanced', optionalName: 'Demo Settings', optionalToolTip: 'Tooltip', outputSeparator: true };
                this.propertyUIOverride.hideProperty('HelpUrl', 'Advanced');
                var content = this.propertyUIOverride.createNewContentAtTop(contentSettings);
                var select = $(content.html('<select id="contosoSelectSPList"></select>')[0].children[0]);
                select.change((e) => {
                    //console.log('change');
                    //var value = (<HTMLSelectElement><any>this).value;
                    var settings: ISetting = { l: (e.currentTarget as HTMLSelectElement).value, v: true };
                    that.propertyUIOverride.setValue('HelpUrl', window.btoa(JSON.stringify(settings)), 'Advanced');
                    that.settings = settings;
                    that.Init.apply(that);
                });
                var html: string[] = [];
                html.push("<option>Typescript</option>");
                html.push("<option>React</option>");
                html.push("<option>WebPack</option>");
                html.push("<option>Office365</option>");
                select[0].innerHTML = html.join("");

                select.val(this.settings.l);
            }
        }

        Init() {
            var that = this;
			 ReactDOM.render(
                <Counter demoSettings={that.settings.l} />,
                that.output
            );
        }
    }
}