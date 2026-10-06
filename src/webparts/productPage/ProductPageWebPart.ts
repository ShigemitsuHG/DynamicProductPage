import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneLabel,
  // PropertyPaneChoiceGroup,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ProductPageWebPartStrings';
import ProductPage from './components/ProductPage';
import { IProductPageProps } from './components/IProductPageProps';



export interface IProductPageWebPartProps {
  description: string;
  fields: string;
  initialProductCode: string;
  showProductName: boolean;
  showProductDesc: boolean;
  showFeatures: boolean;
  showMaker_voice: boolean;
  textMaker_voice:string;
  showPrice: boolean;
  showShops: boolean;
  showExDate: boolean;
  showAllergy: boolean;
  showAlcohol: boolean;
  showAlcoholName: boolean;
  showStorage_type: boolean;
  showPicture: boolean;
  showCakeCut: boolean;
  showTopping: boolean;
  showQuestion: boolean;
  showPoint: boolean;
  showMaterials: boolean;
  showGlossary: boolean;
  showRemoved: boolean;
  showShiage: boolean;
  listName_Main: string;
  listName_Material: string;
  listName_Master: string;
  listName_Glossary: string;
  listName_FAQ: string;
}



// groupFields: [
//   PropertyPaneTextField('description', {
//     label: "説明"
//   }),
//   PropertyPaneTextField('fields', {
//     label: "表示する項目（カンマ区切り）"
//   })
// ]

export default class ProductPageWebPart extends BaseClientSideWebPart<IProductPageWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IProductPageProps> = React.createElement(
      ProductPage,
      {
        context: this.context,
        description: this.properties.description,
        fields: this.properties.fields,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        initialProductCode: this.properties.initialProductCode,
        showProductName: this.properties.showProductName,
        showProductDesc: this.properties.showProductDesc,
        showFeatures: this.properties.showFeatures,
        showMaker_voice: this.properties.showMaker_voice,
        textMaker_voice: this.properties.textMaker_voice,
        showPrice: this.properties.showPrice,
        showShops: this.properties.showShops,
        showExDate: this.properties.showExDate,
        showAllergy: this.properties.showAllergy,
        showAlcohol: this.properties.showAlcohol,
        showAlcoholName: this.properties.showAlcoholName,
        showStorage_type: this.properties.showStorage_type,
        showPicture: this.properties.showPicture,
        showCakeCut: this.properties.showCakeCut,
        showTopping: this.properties.showTopping,
        showQuestion: this.properties.showQuestion,
        showPoint: this.properties.showPoint,
        showMaterials: this.properties.showMaterials,
        showGlossary: this.properties.showGlossary,
        showRemoved: this.properties.showRemoved,
        showShiage: this.properties.showShiage,
        listName_Main: this.properties.listName_Main,
        listName_Material: this.properties.listName_Material,
        listName_Master: this.properties.listName_Master,
        listName_Glossary: this.properties.listName_Glossary,
        listName_FAQ: this.properties.listName_FAQ
      }
    );
    
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    if (this.properties.initialProductCode === undefined) {
      this.properties.initialProductCode = '10100001';
    }
    if (this.properties.showProductName === undefined) {this.properties.showProductName = true}
    if (this.properties.showProductDesc === undefined) {this.properties.showProductDesc = false}
    if (this.properties.showFeatures === undefined) {this.properties.showFeatures = true}
    if (this.properties.showMaker_voice === undefined) {this.properties.showMaker_voice = true}
    if (this.properties.textMaker_voice === undefined) {this.properties.textMaker_voice ="製造者の声"}
    if (this.properties.showPrice === undefined) {this.properties.showPrice = true}
    if (this.properties.showShops === undefined) {this.properties.showShops = true}
    if (this.properties.showExDate === undefined) {this.properties.showExDate = true}
    if (this.properties.showAllergy === undefined) {this.properties.showAllergy = true}
    if (this.properties.showAlcohol === undefined) {this.properties.showAlcohol = true}
    if (this.properties.showAlcoholName === undefined) {this.properties.showAlcoholName = false}
    if (this.properties.showStorage_type === undefined) {this.properties.showStorage_type = true}
    if (this.properties.showPicture === undefined) {this.properties.showPicture = true}
    if (this.properties.showCakeCut === undefined) {this.properties.showCakeCut = true}
    if (this.properties.showTopping === undefined) {this.properties.showTopping = true}
    if (this.properties.showQuestion === undefined) {this.properties.showQuestion = true}
    if (this.properties.showPoint === undefined) {this.properties.showPoint = true}
    if (this.properties.showMaterials === undefined) {this.properties.showMaterials = true}
    if (this.properties.showGlossary === undefined) {this.properties.showGlossary = true}
    if (this.properties.showRemoved === undefined) {this.properties.showRemoved = false}
    if (this.properties.showShiage === undefined) {this.properties.showShiage = false}
    if (this.properties.listName_Main === undefined){this.properties.listName_Main = "01:原材料リスト_ケーキ"}
    if (this.properties.listName_Material === undefined){this.properties.listName_Material = "01:材料マスタ_ケーキ"}
    if (this.properties.listName_Master === undefined){this.properties.listName_Master = "商品マスタ"}
    if (this.properties.listName_Glossary === undefined){this.properties.listName_Glossary = "用語解説と材料特徴"}
    if (this.properties.listName_FAQ === undefined){this.properties.listName_FAQ = "FAQ"}
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: '設定',
              groupFields:[
                PropertyPaneTextField('initialProductCode', {
                  label: "商品コード",
                  description: 'パラメータにもコードを設定できます「例:productCode=10100001」',
                  placeholder: '10100001'
                })
              ]
            },
            {
              groupName: '表示項目の設定',
              groupFields: [
                PropertyPaneCheckbox('showProductName', {text: '商品名'}),
                PropertyPaneCheckbox('showProductDesc', {text: '商品説明*'}),
                PropertyPaneCheckbox('showFeatures', {text: '商品特徴*'}),
                PropertyPaneCheckbox('showMaker_voice', {text: '製造者の声*'}),
                PropertyPaneTextField('textMaker_voice', {placeholder: '製造者の声'}),
                PropertyPaneCheckbox('showPrice', {text: '価格*'}),
                PropertyPaneCheckbox('showShops', {text: '取扱店舗*'}),
                PropertyPaneCheckbox('showExDate', {text: '消費期限*'}),
                PropertyPaneCheckbox('showAllergy', {text: '特定原材料'}),
                PropertyPaneCheckbox('showAlcohol', {text: 'アルコール含有量'}),
                PropertyPaneCheckbox('showAlcoholName', {text: '洋酒の種類'}),
                PropertyPaneCheckbox('showStorage_type', {text: '保存方法*'}),
                PropertyPaneCheckbox('showPicture', {text: '画像'}),
                PropertyPaneCheckbox('showCakeCut', {text: 'ケーキカットのコツ'}),
                PropertyPaneCheckbox('showTopping', {text: 'トッピング'}),
                PropertyPaneCheckbox('showRemoved', {text: '抜ける食材'}),
                PropertyPaneCheckbox('showShiage', {text: '仕上げ'}),
                PropertyPaneCheckbox('showQuestion', {text: 'FAQ'}),
                PropertyPaneCheckbox('showPoint', {text: 'Point'}),
                PropertyPaneCheckbox('showMaterials', {text: '原材料リスト'}),
                PropertyPaneCheckbox('showGlossary', {text: '用語解説'}),
                PropertyPaneLabel('note', {text: '※はデータがある時のみ表示'})
                // PropertyPaneChoiceGroup('activelist', {
                //   label: '▼使用するリスト',
                //   options: [
                //     { key: '01', text: '[01]ケーキ・ギフト・ドルチェ'},
                //     { key: '02', text: '[02]料理'}
                //   ]
                // }),
                // PropertyPaneLabel('', {
                //   text: '　・原材料リスト【' + getListName(this.properties.activelist)[0] + '】'
                // }),
                // PropertyPaneLabel('', {
                //   text: '　・材料マスタ　【' + getListName(this.properties.activelist)[1] + '】'
                // })
              ]
            },
            {
              groupName: '使用リストの設定',
              groupFields: [
                PropertyPaneLabel('', {text: '▼商品マスタ'}),
                PropertyPaneTextField('listName_Master', {
                  // label: '商品マスタ',
                  // description: 'Name,Material[参照],Product[複数参照]',
                  placeholder:'リスト名を指定'
                }),
                PropertyPaneLabel('', {text: '▼原材料リスト'}),
                PropertyPaneTextField('listName_Main', {
                  label: '原材料リスト',
                  description: 'Name,Material[参照],Product[複数参照]',
                  placeholder:'リスト名を指定'
                }),
                PropertyPaneTextField('listName_Material', {
                  label: '材料マスタ',
                  description: 'Name,OriginOrProcessingPlace',
                  placeholder:'リスト名を指定'
                }),
                PropertyPaneLabel('', {text: '▼用語解説'}),
                PropertyPaneTextField('listName_Glossary', {
                  // label: '用語解説',
                  // description: 'Name,Material[参照],Product[複数参照]',
                  placeholder:'リスト名を指定'
                }),
                PropertyPaneLabel('', {text: '▼FAQ'}),
                PropertyPaneTextField('listName_FAQ', {
                  placeholder:'リスト名を指定'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
