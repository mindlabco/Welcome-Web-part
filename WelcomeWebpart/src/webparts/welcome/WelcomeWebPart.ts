import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './WelcomeWebPart.module.scss';
import * as strings from 'WelcomeWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';
import * as pnp from "sp-pnp-js";

export interface IWelcomeWebPartProps {
  description: string;
}

var mythis;
export default class WelcomeWebPart extends BaseClientSideWebPart<IWelcomeWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context
      });
    });
  }

  public constructor() {
    super();
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');

    SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js', { globalExportsName: 'jQuery' }).then((jQuery: any): void => {
      SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js', { globalExportsName: 'jQuery' }).then((): void => {
      });
    });
    //Custom CSS
    SPComponentLoader.loadCss('https://msletb.sharepoint.com/sites/FET/_catalogs/masterpage/School/css/style.css');

  }

  public getDataFromList(): void {

    pnp.sp.web.lists.getByTitle('WelcomeList').items.top(1).orderBy('Modified', false).select('OData__x006f_gt8/Name', '*').expand('OData__x006f_gt8').get().then(function (result) {
      //console.log(result);
      var title = result[0].Title?result[0].Title:"";
      var msg = result[0].Message?result[0].Message:"";
      var ownerID = result[0].OData__x006f_gt8Id?result[0].OData__x006f_gt8Id:"";
      var name = result[0].OData__x006f_gt8.Name?result[0].OData__x006f_gt8.Name:""
      mythis.getUserInfo(name);

      var welcomeTitle = document.getElementById('welTitle');
      var welcomeMsg = document.getElementById('welMsg');
      welcomeTitle.innerHTML = title;
      welcomeMsg.innerHTML = msg;

      //mythis.displayData(result);
    }, function (er) {
      alert("Oops, Something went wrong, Please try after sometime");
      console.log("Error:" + er);
    });
  }

  public getUserInfo(userName): any {
    pnp.sp.profiles.getPropertiesFor(userName).then(function (result) {
      //console.log(result);

      var usEmail = result.Email?result.Email:"";
      document.getElementById('userEmail').setAttribute('href', 'mailto:' + usEmail);
      document.getElementById('userEmail').innerText = usEmail;

      var webURL=mythis.context.pageContext.web.absoluteUrl;
      var tmp1=webURL.split('.com');
      var rootSite = tmp1[0]+'.com'
      var userprofile= rootSite+"/_layouts/15/me.aspx/?p="+usEmail+"&v=work";
      document.getElementById('profilePage').setAttribute('href',userprofile);

      var ownerName = result.DisplayName?result.DisplayName:"";
      document.getElementById('welOwner').innerText = ownerName;

      //result.UserProfileProperties.forEach(function (val) {
        //if (val.Key == "PictureURL") {
          var profileImg = rootSite+"/_layouts/15/userphoto.aspx?size=L&accountname="+usEmail;
          var profileImage = document.getElementById('userImg');
          profileImage.setAttribute('src', profileImg);
        //}
      //})

    })
  }

  public render(): void {
    // Assigning context of the class
    mythis = this;
    this.domElement.innerHTML =
      '<div class="row" style="margin:0px;">' +
      '<div class="col-sm-2">' +
      '<a id="profilePage" href="#" target="_blank"><img style="margin: auto;" id="userImg" src="" class="img-responsive"></a>' +
      '<a id="userEmail" href="#" target="_top"></a>' +
      '</div>' +
      '<div class="col-sm-10">' +
      '<h1 id="welTitle" style="margin-top: 0px;color:#058c90;"></h1>' +
      '<p id="welMsg" style="color: #058c90;"></p>' +
      '<p id="welOwner" style="text-align:right;color:#444;font-weight:700"></p>' +
      '</div>' +
      '</div>';

    //Get data from list
    this.getDataFromList()
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
