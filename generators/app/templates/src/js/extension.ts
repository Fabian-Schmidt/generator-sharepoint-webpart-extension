/// <reference path="extension.d.ts" />

module COMPANY.ExtensionName {

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

    export interface ISharePointSecurityGroupMember {
        Email: string;
        Id: number;
        IsHiddenInUI: boolean;
        IsShareByEmailGuestUser: boolean;
        IsSiteAdmin: boolean;
        LoginName: string;
        PrincipalType: number;
        Title: string;
    }

    export class ExtensionName {
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
            //if (!this.settings.v) {
            //    this.output.innerText += 'Please edit the web part to configure settings for it in the property ui.';
            //}
            {
                this.LoadSharePointData();
                $.when(this.memberGroupsPromise, this.ownerGroupsPromise).then(function(members: ISharePointSecurityGroupMember[], owners: ISharePointSecurityGroupMember[]) {
                    var list = $('<ul></ul');
                    members.forEach(member => {
                        list.append('<li>' + that.RenderUserPresence(member) + '</li>');
                    });
                    owners.forEach(member => {
                        list.append('<li>' + that.RenderUserPresence(member) + '</li>');
                    });
                    that.output.appendChild(list[0]);
                });
            }
        }

        LoadSharePointData() {
            var that = this;

            var memberGroupsDeferred = $.Deferred<ISharePointSecurityGroupMember[]>();
            this.memberGroupsPromise = memberGroupsDeferred.promise();
            var restUrl = _spPageContextInfo.webServerRelativeUrl + '/_api/web/AssociatedMemberGroup/Users';
            $.ajax({ url: restUrl, contentType: "application/json", dataType: "json", headers: { Accept: "application/json; odata=verbose" } }).then(
                function(data) {
                    if (data.value)
                        memberGroupsDeferred.resolve(data.value);
                    if (data.d && data.d.results)
                        memberGroupsDeferred.resolve(data.d.results);
                }, function(xhr, textStatus) {
                    console && console.log('error loading SharePoint member group:' + textStatus);
                    that.output.innerText += 'error loading SharePoint member group:' + textStatus;
                    //xhr.responseJSON['odata.error'].message.value;
                });

            var ownerGroupsDeferred = $.Deferred<ISharePointSecurityGroupMember[]>();
            this.ownerGroupsPromise = ownerGroupsDeferred.promise();
            var restUrl = _spPageContextInfo.webServerRelativeUrl + '/_api/web/AssociatedOwnerGroup/Users';
            $.ajax({ url: restUrl, contentType: "application/json", dataType: "json", headers: { Accept: "application/json; odata=verbose" } }).then(
                function(data) {
                    if (data.value)
                        ownerGroupsDeferred.resolve(data.value);
                    if (data.d && data.d.results)
                        ownerGroupsDeferred.resolve(data.d.results);
                }, function(xhr, textStatus) {
                    console && console.log('error loading SharePoint admin groups:' + textStatus);
                    that.output.innerText += 'error loading SharePoint admin groups:' + textStatus;
                    //xhr.responseJSON['odata.error'].message.value;
                });
        }

        RenderUserPresence(user: ISharePointSecurityGroupMember): string {
            var template = '<nobr>';
            template += ' <span class="ms-imnSpan">';
            template += '  <a href="#" onclick="IMNImageOnClick(event); return false;" class="ms-imnlink ms-spimn-presenceLink" tabindex="-1">';
            template += '   <span class="ms-spimn-presenceWrapper ms-imnImg ms-spimn-imgSize-10x10">';
            template += '    <img title="" alt="No presence information" name="imnmark" class="ms-spimn-img ms-spimn-presence-disconnected-10x10x32" showofflinepawn="1" src="/_layouts/15/images/spimn.png" sip="' + user.Email + '" id="imn0,type=sip" data-themekey="#">';
            template += '   </span>';
            template += '  </a>';
            template += ' </span>';
            template += ' <span class="ms-noWrap ms-imnSpan">';
            template += '  <a href="#" onclick="IMNImageOnClick(event); return false;" class="ms-imnlink" tabindex="-1"><img title="" alt="No presence information" name="imnmark" class="ms-hide" showofflinepawn="1" src="/_layouts/15/images/spimn.png" sip="' + user.Email + '" id="imn1,type=sip" data-themekey="#"></a>';
            template += user.Title;
            template += ' </span>';
            template += '</nobr>';
            return template;
        }
    }
}
