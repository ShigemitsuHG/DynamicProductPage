import * as React from 'react';
import type { IProductPageProps } from './IProductPageProps';
// import { FontSizes } from '@fluentui/react';
// import { useState, useEffect } from "react"
// import strings from 'ProductPageWebPartStrings';

import { spfi, SPFx } from "@pnp/sp"
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
// import { Item } from '@pnp/sp/items';
// import { Fields } from '@pnp/sp/fields/types';



interface IProductPageState {
  product: IProduct | null;
  grouped: GroupedItem[];
  glossaryItems: any;
  faqItems: IFaq[];
}
export interface IProduct {
  Id: number;
  ProductCode: string;
  ProductName: string;
  AlcoholLv: number;
  AllergyList: string[];
  Features: string;
  Maker_voice: string;
  Price: string;
  Shops: string;
  ExDate: Date;
  Alcohol: string;
  Storage_type: string;
  CakeCut: string;
  Toping: string;
  Point: string;
  ProductImage: string;//リストのオリジナルImage列
  AttachmentFiles: IAttachmentFile[];//Image列が格納される添付ファイル
  ProductPictureURL: string;//商品イメージの完全パス
  ProductPage: IHyperlinkField;
  FAQId: number[];
  Removed: string;
  Shiage: string;
  ProductDescTitle: string;
  ProductDesc: string;
}
type GroupedItem = {
  semiProductName: string;
  materials: Item_Main[];
  sortName: string;
};
export interface Item_Main{
  Name: string;//リストから
  MaterialId: number;//リストから
  ProductCodeId : number[];//リストから
  SortName : string;//リストから
  MaterialName: string;
  Origin: string;
  SemiProduct:boolean;
  // ProductCode: ProductCodeLookup[];
}
export interface Item_Material{
  ID: number;//リストから
  Name: string;//リストから
  OriginOrProcessingPlace: string//リストから
  SemiProduct: boolean;
}
export interface Item_Glossary{
  ProductCodeId : number[];//リストから
}
interface IAttachmentFile {
  FileName: string;
  ServerRelativeUrl: string;
}
interface IHyperlinkField {
  Url: string;
  Description: string;
}
interface IFaq {
  Id : number;
  Question : string;
  Answer : string;
}
// type Item = {
//   OData__x002e__x002e__x002e_: string;//半製品
//   OData__x6750__x6599_Id: number;//材料
//   MaterialName: string;
//   Origin: string;
//   // ProductCode: ProductCodeLookup[];
//   ProductCodeId : number[];
// };

// type ProductCodeLookup = {
//   Id: number;
//   Title: string;
// };
// type Field = keyof typeof FIELD_CONFIG;

// const FIELD_CONFIG = {
//   Features:{ label: "商品特徴"},
//   Maker_voice:{ label: "製造者の声"},
//   Price: { label: "価格" },
//   Shops: { label: "取扱店舗" },
//   ExDate: { label: "消費期限" },
//   Allergy: { label: "特定原材料等29品目" },
//   Alcohol: { label: "アルコール含有量" },
//   Storage_type: { label: "保存方法" },
//   Picture: { label: "画像" },
//   CakeCut: { label: "ケーキカットのコツ" },
//   Topping: { label: "トッピングの仕方" },
//   Question: { label: "よくある質問" },
//   Point: { label: "Point" },
//   Materials: { label: "原材料リスト" },
//   Glossary: { label: "用語解説・材料特徴について" }
// } as const;


export default class ProductPage extends React.Component<IProductPageProps , IProductPageState> {
 
  // private shouldShow(field: Field): boolean {
  //   return !this.state.fieldSet || this.state.fieldSet.has(field);
  // }

  public render(): React.ReactElement<IProductPageProps> {

    // ✅ URLパラメータ取得
    // const urlParams = new URLSearchParams(window.location.search);
    // const productCode = urlParams.get("productCode");

    // ✅ PowerApps URL生成
    // const productCode = this.state.product?.ProductCode;
    // const url = `https://apps.powerapps.com/play/e/default-c6fe6c6f-1efc-472f-b587-c79eea88fd36/a/eccf8198-d64e-445d-aff3-79a1f4a510d9?tenantId=c6fe6c6f-1efc-472f-b587-c79eea88fd36&hint=286cfcce-92f6-40b7-b40c-7500a8caa583&sourcetime=1781489861589&productCode=${encodeURIComponent(productCode || "")}`;
    let alcoholImage = null;
    if ( this.state.product?.AlcoholLv === 1){
      alcoholImage = "https://shigemitsu365.sharepoint.com/sites/products/Shared%20Documents/%E7%94%BB%E5%83%8F/%E3%82%A2%E3%83%AB%E3%82%B3%E3%83%BC%E3%83%AB%E4%BD%BF%E7%94%A8%E9%87%8F/AlcoholLv1.jpg";    
    } else if (this.state.product?.AlcoholLv === 2) {
      alcoholImage = "https://shigemitsu365.sharepoint.com/sites/products/Shared%20Documents/%E7%94%BB%E5%83%8F/%E3%82%A2%E3%83%AB%E3%82%B3%E3%83%BC%E3%83%AB%E4%BD%BF%E7%94%A8%E9%87%8F/AlcoholLv2.jpg";
    } else if (this.state.product?.AlcoholLv === 3) {
      alcoholImage = "https://shigemitsu365.sharepoint.com/sites/products/Shared%20Documents/%E7%94%BB%E5%83%8F/%E3%82%A2%E3%83%AB%E3%82%B3%E3%83%BC%E3%83%AB%E4%BD%BF%E7%94%A8%E9%87%8F/AlcoholLv3.jpg";
    } else if (this.state.product?.AlcoholLv === 4) {
      alcoholImage = "https://shigemitsu365.sharepoint.com/sites/products/Shared%20Documents/%E7%94%BB%E5%83%8F/%E3%82%A2%E3%83%AB%E3%82%B3%E3%83%BC%E3%83%AB%E4%BD%BF%E7%94%A8%E9%87%8F/AlcoholLv4.jpg";
    }
    document.title=this.state.product?.ProductName ?? "";
    return (

      <div style={{width: "100%"}}>
        
        <div style={{ display: "flex"}}>
          <div style={{ width: "50%" ,paddingRight:"30px"}}>
            <h2 style={{fontSize:"24px",fontWeight: "bold"}}>
              {
                this.props.showProductName &&(
                  this.state.product?.ProductName
                )
              }
            </h2>
            {this.props.showProductDesc && this.state.product?.ProductDesc && (
              <>
                <h2 style={{ fontSize: "18px", fontWeight: "bold" }}>
                    {this.state.product?.ProductDescTitle ?? "商品説明"}：
                </h2>
                <pre style={{ fontSize: "18px", whiteSpace: "pre-line" }}>
                  {this.state.product?.ProductDesc}
                </pre>
              </>
            )}            
            {this.props.showFeatures && this.state.product?.Features && (
              <>
                <h2 style={{ fontSize: "18px", fontWeight: "bold" }}>
                  商品特徴：
                </h2>
                <pre style={{ fontSize: "18px", whiteSpace: "pre-line" }}>
                  {this.state.product?.Features}
                </pre>
              </>
            )}
            {this.props.showMaker_voice && this.state.product?.Maker_voice && (
              <>
                <h2 style={{ fontSize: "18px", fontWeight: "bold" }}>
                  {this.props.textMaker_voice}：
                </h2>
                <pre style={{ fontSize: "18px", whiteSpace: "pre-line" }}>
                  {this.state.product?.Maker_voice}
                </pre>
              </>
            )}
            {this.props.showPrice && this.state.product?.Price &&(
              <>
                <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  価格：
                </h2>
                <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line" }}>
                  {this.state.product?.Price}
                </pre>
              </>
            )}
            {this.props.showShops && this.state.product?.Shops &&(
              <>
                <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  取扱店舗：
                </h2>
                <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line" }}>
                  {this.state.product?.Shops}
                </pre>
                {/* <h2>
                  <span style={{ fontSize: "18px", fontWeight: "bold" }}>
                    取扱店舗：
                  </span>
                  <span style={{ fontSize: "18px", fontWeight: "normal", marginLeft: "8px" }}>
                    {this.state.product?.Shops}
                  </span>
                </h2> */}
              </>
            )}
            {this.props.showExDate && this.state.product?.ExDate && (
              <>
                <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  消費期限：
                </h2>
                <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line" }}>
                  {this.state.product?.ExDate}
                </pre>
                {/* <h2>
                  <span style={{ fontSize: "18px", fontWeight: "bold" }}>
                    消費期限：
                  </span>
                  <span style={{ fontSize: "18px", fontWeight: "normal", marginLeft: "8px" }}>
                    {this.state.product?.ExDate}
                  </span>
                </h2> */}
              </>
            )}
            {this.props.showAllergy && (
              <>
                <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  特定原材料等29品目：
                  {/* <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line"}}>
                    {this.state.product?.Allergy}
                  </pre> */}
                 <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line"}}>
                    {this.state.product?.AllergyList?.join("・")}
                  </pre>
                </h2>
              </>
            )}
            {this.props.showAlcohol && (
              <>
                <h2>
                  <span style={{ fontSize: "18px", fontWeight: "bold" }}>
                    アルコール含有量
                  </span>
                  <span style={{ fontSize: "12px", fontWeight: "nomal",marginLeft: "8px"}}>
                    (洋酒を強く感じる度合い)
                  </span>
                  <span style={{ fontSize: "18px", fontWeight: "bold" }}>
                    ：
                  </span>
                </h2>
                <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                  {alcoholImage ? (
                    <>
                      <img src={alcoholImage} style={{ width: "120px", marginBottom: "10px" }} />
                      <span style={{ fontSize: "14px" }}>
                        洋酒の種類 : {this.state.product?.Alcohol}
                      </span>
                    </>
                  ) : (
                    <span style={{ fontSize: "18px", fontWeight: "normal" }}>
                      不使用
                    </span>
                  )}
                </div>
              </>
            )}
            {this.props.showAlcoholName && (
              <>
                <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  洋酒の種類：
                 <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line"}}>
                    {this.state.product?.Alcohol}
                  </pre>
                </h2>
              </>
            )}
            {this.props.showStorage_type && this.state.product?.Storage_type &&  (
              <>
                <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  保存方法：
                </h2>
                <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line" }}>
                  {this.state.product?.Storage_type}
                </pre>
                {/* <h2>
                  <span style={{ fontSize: "18px", fontWeight: "bold" }}>
                    保存方法：
                  </span>
                  <span style={{ fontSize: "18px", fontWeight: "normal", marginLeft: "8px" }}>
                    {this.state.product?.Storage_type}
                  </span>
                </h2> */}
              </>
            )}
            {this.props.showPicture && (
              <>
                <div>
                  {this.state.product?.ProductPictureURL && (
                    <img src={this.state.product?.ProductPictureURL} style={{ width: "100%" }} />
                  )}
                </div>
              </>
            )}
            {this.props.showCakeCut && (
              <>
                {/* <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  ケーキカットのコツ：
                  <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line"}}>
                    {this.state.product?.CakeCut}
                  </pre>
                </h2> */}
                  <h2 style={{ fontSize: "18px", fontWeight: "bold" }}>
                    ケーキカットのコツ：
                  </h2>
                  <div
                    style={{ fontSize: "18px", fontWeight: "normal" }}
                    dangerouslySetInnerHTML={{
                      __html: this.state.product?.CakeCut || ""
                    }}
                  />
              </>
            )}
            {this.props.showTopping && (
              <>
                <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  トッピングの仕方：
                  {console.log("VideoURL:"+this.state.product?.Toping)}
                  {this.state.product?.Toping && (
                      <video width="100%" controls>
                        <source src={this.state.product?.Toping} type="video/mp4" />
                      </video>
                    )}
                </h2>
              </>
            )}
            {this.props.showRemoved && (
              <>
                <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  抜ける食材：
                </h2>
                <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line" }}>
                  {this.state.product?.Removed}
                </pre>
              </>
            )}
            {this.props.showRemoved && (
              <>
                <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  仕上げ：
                </h2>
                <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line" }}>
                  {this.state.product?.Shiage}
                </pre>
              </>
            )}
            {this.props.showQuestion && (
              <>
                {/* <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  よくある質問：
                  <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line"}}>
                    {this.state.product?.Question}
                  </pre>
                </h2> */}
                <h2 style={{ fontSize: "18px", fontWeight: "bold" }}>
                  よくある質問：
                </h2>

                {this.state.faqItems?.map((faq, index) => (
                  <div key={faq.Id} style={{ marginBottom: "10px" }}>
                    <div style={{ fontSize: "18px", fontWeight: "bold" ,marginLeft: "10px"}}>
                      Q{index + 1}. {faq.Question}
                    </div>

                    <pre
                      style={{
                        fontSize: "18px",
                        fontWeight: "normal",
                        whiteSpace: "pre-line",
                        marginLeft: "20px"
                      }}
                    >
                      {faq.Answer}
                    </pre>
                  </div>
                ))}
              </>
            )}
            {this.props.showPoint && (
              <>
                <h2 style={{fontSize:"18px",fontWeight: "bold"}}>
                  Point：
                  <pre style={{ fontSize: "18px", fontWeight: "normal", whiteSpace: "pre-line"}}>
                    {this.state.product?.Point}
                  </pre>
                </h2>
              </>
            )}
          </div>
          <div style={{ width: "50%" }}>
            {this.props.showMaterials && (
              <>
                <h2 style={{fontSize:"20px",fontWeight: "bold"}}>
                  原材料リスト
                </h2>
                {/* ✅ 見出し */}
                <li
                  style={{
                      display: "grid",
                      gridTemplateColumns: "150px 1fr",
                      borderBottom: "1px solid #ccc",
                      padding: "4px 0 4px 30px"
                  }}
                >
                  <span>材料</span>
                  <span>産地または加工地</span>
                </li>
                {this.state.grouped.map(group => (
                  <div key={group.semiProductName}>
                    <h3 style={{paddingLeft:"20px"}}>
                      {group.semiProductName}
                    </h3>
                    <ul style={{ listStyle: "none", padding: 0,  paddingLeft: "30px" }}>

                      {group.materials.map((item, i) => (
                        <li
                          key={i}
                          style={{
                            display: "grid",
                            gridTemplateColumns: "150px 1fr",
                            borderBottom: "1px solid #ccc",
                            padding: "4px 0"
                          }}
                        >
                          <span style={{fontWeight: item.SemiProduct? "bold":"normal"}}>
                            {item.MaterialName}
                          </span>
                          <span style={{ color: item.SemiProduct ? 'blue' : '#000' }}>
                            {item.Origin}
                          </span>

                        </li>
                      ))}
                    </ul>
                  </div>
                ))}
              </>
            )}
          </div>
        </div>
        <div>
            {this.props.showGlossary && (
              <>
                <h2 style={{fontSize:"20px",fontWeight: "bold"}}>
                  用語解説・材料特徴について
                </h2>
                <ul style={{ listStyle: "none", padding: 0,  paddingLeft: "10px" }}>
                  {/* ✅ 見出し */}
                  <li
                    style={{
                      display: "grid",
                      gridTemplateColumns: "150px 1fr",
                      fontWeight: "bold",
                      borderBottom: "2px solid #aaa",
                      padding: "6px 0"
                    }}
                  >
                    <span>用語</span>
                    <span>解説</span>
                  </li>
                  {this.state.glossaryItems.map((item: any, i: number) => (
                    <li
                      key={i}
                      style={{
                        display: "grid",
                        gridTemplateColumns: "150px 1fr",
                        borderBottom: "1px solid #ccc",
                        padding: "4px 0"
                      }}
                    >
                      <span>{item.term}</span>
                      <span>{item.explanation}</span>
                    </li>
                  ))}
                </ul>
              </>
            )}
        </div>
      </div>
    );
  }
  
  private sp;
  //初期化
  constructor(props: IProductPageProps) {
    super(props);

    this.state = {
      product: null,
      grouped: [],
      glossaryItems:[],
      faqItems:[]
    };
    // ✅ sp初期化
    this.sp = spfi("https://shigemitsu365.sharepoint.com/sites/products")
      .using(SPFx(this.props.context));
  }

  public componentDidMount(): void {

    this.initial();

    // this.props.context.spHttpClient
    //   .get(
    //     // `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('商品マスタ')/items?$filter=ProductCode eq '${productCode}'`,
    //     // SPHttpClient.configurations.v1
    //     `https://shigemitsu365.sharepoint.com/sites/products/_api/web/lists/getbytitle('商品マスタ')/items?$filter=ProductCode eq '${productCode}'`,
    //      SPHttpClient.configurations.v1
    //   )
    //   .then(res => res.json())
    //   .then(data => {
        
    //     console.log(data);
    //     // ✅ データがあればstateにセット
    //     if (data.value.length > 0) {
    //       console.log(data.value[0].Id);
          
    //       this.setState({
    //         product: data.value[0],
    //         fieldSet: fieldSet
    //       });
    //       this.loadData(data.value[0].Id,productCode);
    //       this.loadData2(data.value[0].Id,productCode);
    //     }
    //   });
  }
  private async initial(){
   const params = new URLSearchParams(window.location.search);
    let productCode = params.get("productCode");
    // const urlFields = params.get("fields") ?? "";
    // const propertyFields = this.props.fields ?? "";
    console.log("コード:"+productCode);
    console.log("listName:原材料リスト=[%s] 材料マスタ=[%s] 商品マスタ=[%s] 用語解説=[%s]",  
                              this.props.listName_Main,
                              this.props.listName_Material,
                              this.props.listName_Master,
                              this.props.listName_Glossary);


    if (!productCode) {
      productCode = this.props.initialProductCode;
    }
    console.log("productCode:",productCode);
    

    // console.log("ServerRelativeUrl:",pictureFile[0].AttachmentFiles[0].ServerRelativeUrl);

    const datas :IProduct[] = await this.sp.web.lists
      .getByTitle(this.props.listName_Master)
      .items
      .filter(`ProductCode eq '${productCode}'`)
      .select("*", "AttachmentFiles")
      .expand("AttachmentFiles")
      .top(1)();
    // console.log("Attachment",datas[0].AttachmentFiles);
    // console.log("AttachmentLength",datas.length);
    
    let data :IProduct | null = null;
    if (datas.length > 0) {
      data = datas[0];

      const imageInfo = JSON.parse(data.ProductImage);
      let attachment: IAttachmentFile | undefined;

      for (const file of data.AttachmentFiles) {
        if (file.FileName === imageInfo.fileName) {
          attachment = file;
          const absoluteUrl =`${window.location.origin}${attachment?.ServerRelativeUrl ?? ""}`;
          data.ProductPictureURL=absoluteUrl;
          break;
        }
      }
      console.log("masterData:",data);
      console.log("PictureURL",data.ProductPictureURL);
      console.log("PageLink",data.ProductPage?.Url);
    }
    

    this.setState({
      product: data
    });
    if (data != null) {
      const id = data.Id;
      this.loadData_Material(id,productCode);
      this.loadData_Glossary(id,productCode);
      this.loadData_Faq(data.FAQId);
    }
  }
  private async loadData_Faq(ids: number[]) {
    if (ids.length === 0) {
      return [];
    }
    console.log("faqIds",ids);
    const filter = ids
      .map(id => `ID eq ${id}`)
      .join(" or ");

    const data: IFaq[] = await this.sp.web.lists
      .getByTitle(this.props.listName_FAQ)
      .items
      .filter(filter)
      .top(5000)();
    console.log("faqItems",data);
    this.setState({
      faqItems: data
    });
  }


  //用語解説データセット
  private async loadData_Glossary(id : number, productCode: any){
    console.log("用語解説.start code:",productCode);
    console.log("用語解説ID:", this.state.product?.Id);
    try{
      const data :Item_Glossary[] = await this.sp.web.lists
        .getByTitle(this.props.listName_Glossary)
        .items
        .top(5000)();
      console.log("用語解説全件:",data)
      // ✅ ③ フィルター（複数Lookup）
      const filtered = data.filter(item =>
        // item.ProductCodeId && item.ProductCodeId.indexOf(id) !== -1
        item.ProductCodeId?.indexOf(id) !== -1
      );
      console.log("抽出後:", filtered);
      // ✅ ④ Stateにセット
      this.setState({
        glossaryItems: filtered
      });

    } catch (e) {
      console.error("ERROR:", e);
    }
  }

  //原材料データセット
  //id 商品マスタの商品id(内部id)
  private async loadData_Material(id : number, productCode: any) {
    console.log("id:",id);
    console.log("原材料データ.start code:", productCode);
    console.log("原材料データID:",this.state.product?.Id);

    try {
      const allItemData : Item_Main[]  = await this.sp.web.lists
        .getByTitle(this.props.listName_Main)
        .items
        // .filter(`substringof('${productCode}', ProductCode)`)
        // .filter(`ProductCode eq '${productCode}'`)
        .top(5000)();
      console.log("Item_Main:",allItemData);
      const materials : Item_Material[]  = await this.sp.web.lists
        .getByTitle(this.props.listName_Material)
        .items
        .top(5000)();
      console.log("Item_Material:",materials);
      // console.log("材料取得" , JSON.stringify(materials));
      //Map化（高速化）
      const materialMap: { [key: number]: Item_Material } = {};
        materials.forEach(m => {
          materialMap[m.ID] = m;
      });
      console.log("Map化",materialMap);

      const enriched = allItemData.map(item => {
        const material = materialMap[item.MaterialId];
        // console.log("material:",item.MaterialId,item.Name,  material);
        return {
          ...item, // 元データをコピー
          MaterialName: material?.Name,
          Origin: material?.OriginOrProcessingPlace,
          SemiProduct: material?.SemiProduct
        };
      });
      console.log("enriched:", enriched);
      // ✅ 「含まれている」条件
      const filtered = enriched.filter(item =>
        // item.ProductCode?.some(pc => pc.Title === productCode)
        item.ProductCodeId?.indexOf(id) !== -1
      );
      console.log("id",id);
      console.log("filtered:", filtered);
      //半製品名をキーにしてキー付きリストに代入
      const groupedData = filtered.reduce<Record<string, Item_Main[]>>(
        //acc=配列　item=1件ずつのデータ
        (acc, item) => {
          const key = item.Name || "未分類";
          if (!acc[key]) acc[key] = [];
          acc[key].push(item);
          return acc;
        },
        {}
      );
      console.log("groupedData:", groupedData);
      // 各グループ内を SortName 順に並び替え
      Object.keys(groupedData).forEach(key => {
        groupedData[key].sort((a, b) =>
          (a.SortName || "").localeCompare(b.SortName || "", "ja")
        );
      });
      //キー付きリストからGroupedItem[]の配列に変換
      const result :GroupedItem[] = Object.keys(groupedData)
        // .sort((a:GroupedItem,b:GroupedItem) => a.semiProduct.localeCompare(b.semiProduct))
        .map(key => ({
          semiProductName: key,
          materials: groupedData[key],
          sortName: groupedData[key]?.[0]?.SortName
      }))
      .sort((a:GroupedItem,b:GroupedItem) => a.sortName.localeCompare(b.sortName))
      ;
      console.log("result:", result);
      // ✅ state更新
      this.setState({
        grouped: result
      });
    } catch (e) {
      console.error("ERROR:", e);
    }
  }

}