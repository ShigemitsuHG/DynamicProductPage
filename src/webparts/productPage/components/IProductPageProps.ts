export interface IProductPageProps {
  description: string;
  fields: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  initialProductCode: string;
  showProductName: boolean;
  showProductDesc: boolean;
  showFeatures: boolean;
  showMaker_voice: boolean;
  textMaker_voice: string;
  showPrice: boolean;
  showShops: boolean;
  showExDate: boolean;
  showAllergy: boolean;
  showAlcoholName: boolean;
  showAlcohol: boolean;
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
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IProductPageProps {
  context: WebPartContext;
}

