import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeProvider, ThemeChangedEventArgs } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './LinkCardsWebPart.module.scss';
import * as strings from 'LinkCardsWebPartStrings';

interface ILinkCard { id:string; title:string; description:string; imageUrl:string; pageUrl:string; backgroundColor:string; textColor:string; }
export interface ILinkCardsWebPartProps { cards: ILinkCard[]; }
interface IFileItem { FileRef:string; FileLeafRef:string; }
interface IRestResult<T> { value:T[]; }

export default class LinkCardsWebPart extends BaseClientSideWebPart<ILinkCardsWebPartProps> {
  private _themeProvider: ThemeProvider | undefined;
  private _theme: IReadonlyTheme | undefined;
  private _editing:boolean = false;
  private _assets:IFileItem[]=[];
  private _pages:IFileItem[]=[];

  protected async onInit():Promise<void>{
    this._themeProvider=this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._theme=this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this,this._onThemeChanged);
    if(!Array.isArray(this.properties.cards)){this.properties.cards=[];}
    await this._loadChoices();
  }
  public render():void{
    this._applyTheme();
    const editButton=this.displayMode===2?`<button class="${styles.editButton}" data-action="toggle-edit">${this._editing?'Zavřít nastavení':'Upravit karty'}</button>`:'';
    this.domElement.innerHTML=`<section class="${styles.linkCards}">${editButton}${this._editing?this._renderEditor():''}${this._renderCards()}</section>`;
    this._bindEvents();
  }
  private _renderCards():string{
    if(this.properties.cards.length===0){return `<div class="${styles.empty}">Zatím nejsou přidané žádné karty. V režimu úprav zvolte „Upravit karty“.</div>`;}
    return `<div class="${styles.grid}">${this.properties.cards.map((c:ILinkCard):string=>`<a class="${styles.card}" href="${this._e(c.pageUrl||'#')}" style="--lc-card-bg:${this._e(c.backgroundColor||'#ffffff')};--lc-card-text:${this._e(c.textColor||'#242424')}" aria-label="${this._e(c.title)}"><div class="${styles.imageArea}">${c.imageUrl?`<img src="${this._e(c.imageUrl)}" alt="" />`:''}</div><div class="${styles.content}"><h3 class="${styles.title}">${this._e(c.title)}</h3><p class="${styles.description}">${this._e(c.description)}</p></div></a>`).join('')}</div>`;
  }
  private _renderEditor():string{
    const rows=this.properties.cards.map((c:ILinkCard,i:number):string=>`<div class="${styles.editorCard}" data-index="${i}"><div class="${styles.actions}"><button data-action="up" title="Posunout nahoru">↑</button><button data-action="down" title="Posunout dolů">↓</button></div><label>Nadpis<input data-field="title" value="${this._e(c.title)}"></label><label>Popis<textarea data-field="description">${this._e(c.description)}</textarea></label><label>Barva karty<input data-field="backgroundColor" type="color" value="${this._e(c.backgroundColor||'#ffffff')}"></label><div class="${styles.actions}"><button data-action="image">Obrázek</button><button data-action="page">Stránka</button><button data-action="remove">Odebrat</button></div><input data-field="imageUrl" type="hidden" value="${this._e(c.imageUrl)}"><input data-field="pageUrl" type="hidden" value="${this._e(c.pageUrl)}"></div>`).join('');
    return `<div class="${styles.editor}"><strong>Nastavení karet</strong>${rows}<div class="${styles.editorFooter}"><button data-action="add">Přidat kartu</button><button data-action="save">Uložit změny</button></div></div>`;
  }
  private _bindEvents():void{
    this.domElement.querySelectorAll<HTMLElement>('[data-action]').forEach((el:HTMLElement):void=>el.addEventListener('click',(ev:MouseEvent):void=>{ev.preventDefault();this._handleAction(el);}));
  }
  private _handleAction(el:HTMLElement):void{
    const action=el.dataset.action;
    if(action==='toggle-edit'){this._editing=!this._editing;this.render();return;}
    if(action==='add'){this.properties.cards.push({id:this._id(),title:'Nová karta',description:'Popis karty',imageUrl:'',pageUrl:'',backgroundColor:'#ffffff',textColor:'#242424'});this.render();return;}
    if(action==='save'){this._readEditor();this.context.propertyPane.refresh();this.render();return;}
    const row=el.closest<HTMLElement>('[data-index]'); if(!row){return;} const index=Number(row.dataset.index); this._readEditor();
    if(action==='remove'){this.properties.cards.splice(index,1);}
    if(action==='up'&&index>0){[this.properties.cards[index-1],this.properties.cards[index]]=[this.properties.cards[index],this.properties.cards[index-1]];}
    if(action==='down'&&index<this.properties.cards.length-1){[this.properties.cards[index+1],this.properties.cards[index]]=[this.properties.cards[index],this.properties.cards[index+1]];}
    if(action==='image'){this._chooseFile(index,this._assets,'Vyberte obrázek',true);return;}
    if(action==='page'){this._chooseFile(index,this._pages,'Vyberte cílovou stránku',false);return;}
    this.render();
  }
  private _readEditor():void{
    this.domElement.querySelectorAll<HTMLElement>('[data-index]').forEach((row:HTMLElement):void=>{const i=Number(row.dataset.index);const card=this.properties.cards[i];if(!card){return;}row.querySelectorAll<HTMLInputElement|HTMLTextAreaElement>('[data-field]').forEach((input):void=>{const f=input.dataset.field;if(f==='title'||f==='description'||f==='imageUrl'||f==='pageUrl'||f==='backgroundColor'){card[f]=input.value;}});});
  }
  private _chooseFile(index:number,items:IFileItem[],title:string,isImage:boolean):void{
    const labels=items.slice(0,100).map((x:IFileItem,i:number):string=>`${i+1}: ${x.FileLeafRef}`).join('\n');
    const selected=window.prompt(`${title}:\n${labels}\n\nZadejte číslo položky:`);
    if(!selected){return;}const item=items[Number(selected)-1];if(!item){window.alert('Neplatná volba.');return;}
    if(isImage){this.properties.cards[index].imageUrl=item.FileRef;}else{this.properties.cards[index].pageUrl=item.FileRef;}this.render();
  }
  private async _loadChoices():Promise<void>{
    const web=this.context.pageContext.web.absoluteUrl;
    const assetUrl=`${web}/_api/web/lists/getbytitle('Site Assets')/items?$select=FileRef,FileLeafRef&$filter=FSObjType eq 0&$orderby=FileLeafRef`;
    const pageUrl=`${web}/_api/web/lists/getbytitle('Site Pages')/items?$select=FileRef,FileLeafRef&$filter=FSObjType eq 0&$orderby=FileLeafRef`;
    this._assets=await this._getFiles(assetUrl).catch(():IFileItem[]=>[]);
    this._pages=await this._getFiles(pageUrl).catch(():IFileItem[]=>[]);
  }
  private async _getFiles(url:string):Promise<IFileItem[]>{const response:SPHttpClientResponse=await this.context.spHttpClient.get(url,SPHttpClient.configurations.v1,{headers:{accept:'application/json;odata.metadata=none'}});if(!response.ok){throw new Error(response.statusText);}const data:IRestResult<IFileItem>=await response.json() as IRestResult<IFileItem>;return data.value;}
  private _applyTheme():void{const t:IReadonlyTheme|undefined=this._theme;const s=this.domElement.style;s.setProperty('--lc-text',t?.semanticColors?.bodyText||'#242424');s.setProperty('--lc-subtext',t?.semanticColors?.bodySubtext||'#605e5c');s.setProperty('--lc-bg',t?.semanticColors?.bodyBackground||'#ffffff');s.setProperty('--lc-divider',t?.semanticColors?.bodyDivider||'#edebe9');s.setProperty('--lc-primary',t?.palette?.themePrimary||'#0078d4');s.setProperty('--lc-focus',t?.semanticColors?.focusBorder||t?.palette?.themePrimary||'#0078d4');}
  private _onThemeChanged=(args:ThemeChangedEventArgs):void=>{this._theme=args.theme;this.render();};
  private _e(value:string):string{return value.replace(/[&<>'"]/g,(c:string):string=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]||c));}
  private _id():string{return `${Date.now()}-${Math.random().toString(16).slice(2)}`;}
  protected get dataVersion():Version{return Version.parse('1.0');}
  protected getPropertyPaneConfiguration():IPropertyPaneConfiguration{return{pages:[{header:{description:strings.PropertyPaneDescription},groups:[{groupName:strings.BasicGroupName,groupFields:[PropertyPaneLabel('editorInfo',{text:strings.EditCardsLabel})]}]}]};}
  protected onDispose():void{this._themeProvider?.themeChangedEvent.remove(this,this._onThemeChanged);}
}
