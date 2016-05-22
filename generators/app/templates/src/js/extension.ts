/// <reference path="extension.d.ts" />

module COMPANY.<%= name %> {

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

    export class <%= name %> {
        propertyUIOverride: AppPartPropertyUIOverride.AppPartPropertyUIOverride = null;
        output: HTMLDivElement;
        webPartId: string;
        settings: ISetting;

        constructor(webPartDOMObject: HTMLDivElement) {
            this.output = <HTMLDivElement>$('<div></div>')[0];
            $(webPartDOMObject).after(this.output);
            if (webPartDOMObject.parentElement.attributes[<any>'webpartid2'])
                this.webPartId = webPartDOMObject.parentElement.attributes[<any>'webpartid2'].value;
            else
                this.webPartId = webPartDOMObject.parentElement.attributes[<any>'webpartid'].value;
            try {
                this.settings = JSON.parse(window.atob(webPartDOMObject.parentElement.attributes[<any>'helplink'].value));
            } catch (err) {
                //init default settings
                this.settings = { v: false };
            }
            this.propertyUIOverride = new AppPartPropertyUIOverride.AppPartPropertyUIOverride(this.webPartId);
            this.propertyUIOverride.IsActive().done(this.PropertyUIOverrideActive.bind(this));

            this.Init();
        }
        PropertyUIOverrideActive(isActive: boolean) {
            if (isActive) {
                var contentSettings = { category: 'Advanced', optionalName: 'Bar', optionalToolTip: 'Tooltip', outputSeparator: true };
                this.propertyUIOverride.hideProperty('HelpUrl', 'Advanced');
                var content = this.propertyUIOverride.createNewContentAtTop(contentSettings);
                var select = $(content.html('<select id="contosoSelectSPList"></select>')[0].children[0]);
                select.change((e) => {
                    //console.log('change');
                    //var value = (<HTMLSelectElement><any>this).value;
                    var settings: ISetting = { l: (<HTMLSelectElement>e.currentTarget).value, v: true };
                    this.propertyUIOverride.setValue('HelpUrl', window.btoa(JSON.stringify(settings)), 'Advanced');
                });
                var html: string[] = [];
                html.push("<option>Foo</option>");
                html.push("<option>FooBar</option>");
                html.push("<option>Bar</option>");
                html.push("<option>Office365</option>");
                select[0].innerHTML = html.join("");

                select.val(this.settings.l);
            }
        }
        memberGroupsPromise: JQueryPromise<ISharePointSecurityGroupMember[]>;
        ownerGroupsPromise: JQueryPromise<ISharePointSecurityGroupMember[]>;

        Init() {
            var that = this;
			
			that.output.innerText += 'Web Part Loaded! Settings: ' + this.settings.l;
        }
    }
}
