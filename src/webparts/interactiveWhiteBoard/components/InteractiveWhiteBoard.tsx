import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as React from 'react';
import styles from './InteractiveWhiteBoard.module.scss';
import { IInteractiveWhiteBoardProps } from './IInteractiveWhiteBoardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IInteractiveWhiteBoardState } from './IInteractiveWhiteBoardState';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { MessageBar, MessageBarType, IStackProps, Stack, DefaultButton, IconButton } from 'office-ui-fabric-react';  
import * as jquery from 'jquery';
import { TextField } from 'office-ui-fabric-react';
require ('../assets/carousel.css');
require ('../assets/InteractiveWebPart.css');
import * as $ from "jquery";
import "slick-carousel";

 
import "@pnp/sp/webs";  
import "@pnp/sp/lists";  
import "@pnp/sp/items"; 
import { Form, sp } from "@pnp/sp/presets/all";
import { Icon } from 'office-ui-fabric-react/lib/Icon';


import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardDetails,
  DocumentCardTitle,
  IDocumentCardPreviewProps,
  DocumentCardLocation,
  DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';



import {
  PrimaryButton,
  autobind,
  Dropdown,
  IDropdown,
  ITextFieldStyles,
  ImageFit
} from 'office-ui-fabric-react';
import { ITitleItem } from './ITitleItem';
import { IRecognitionHistoryItem } from "./IRecognitionHistoryItem";
import { Carousel, CarouselButtonsDisplay, CarouselButtonsLocation, CarouselIndicatorShape, Placeholder } from "@pnp/spfx-controls-react";
import { useState,  Dispatch, SetStateAction } from "react";


const verticalStackProps: IStackProps = {  
  styles: { root: { overflow: 'hidden', width: '100%' } },  
  tokens: { childrenGap: 20 }  
};  



// const[state, setstate]= useState("red");
const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 300 } };

// function   ButtonChangeColor(){
 

//    function myFun(){
//     if(state==="red"){
//       setstate("blue");
//     }
//     else{
//       setstate("red")
//     }
//    }
//   }
export default class InteractiveWhiteBoard extends React.Component<IInteractiveWhiteBoardProps, IInteractiveWhiteBoardState> {
  
  recognisationItem: any;
  TitleItem: any;
  static props: any;
  public TitleName:any;

 
  constructor(props: IInteractiveWhiteBoardProps, state: IInteractiveWhiteBoardState) {
    super(props);  
   
   
    this.state = {
      items:[{
        Title:[],
        child:[],
      }],
      initialPage:true,

      clicked: false,
      recognitionmsg:'',
      errorMessage:'',
      status: 'Ready',
      columns:[],
      history_1:[],
      history_1_filter:[],

      IconItem:[],
      TitleItems: [],
      RecognisationItems:[],
      carouselItems:[],
      carouselJSXItems:null,
      showForm: false,
      showItem: false,
      Received: false,
      Sent:false,
      isEditFormOpen: false,
       buttonSet:{
        Title:"",
        Icon:" "
       },
     
      TitleItem:[],
      recognisationItem:{
        To:"",
        Recognition_Titles:"",
        Appreciation_Message:"",
        Title:" ",
        
        From:"",
        ImageURL:"",
        Subject:""
      },

      title: '',  
      users: {
        id:"",
        Name:"",
       emailId:"",
       resultdata:[]
      },  
      showMessageBar: false,
      disabledbtns:true,
      selectedbuttonId:"",
      selectbutton:false 
     };

     sp.setup({  
      spfxContext : this.props.context  
    });  
  
  

     InteractiveWhiteBoard.onclickofButton=InteractiveWhiteBoard.onclickofButton.bind(this);
     //this.onclickofButton=this.onclickofButton.bind(this);
     this.getCarouselItems=this.getCarouselItems.bind(this);
    // this.showItem = this.showItem.bind(this);
    this.showform = this.showform.bind(this);
    this.initialpage=this.initialpage.bind(this);
    this.onchangeofAM =this.onchangeofAM.bind(this);
    this.getTitleInformation=this.getTitleInformation.bind(this);
    this.myFun =this.myFun.bind(this);
  
}




  public async sendEmail() {
    this.RetrieveList.bind(this);
    debugger;
    const appreciationPreview=document.getElementById("recognitionId");
    
      console.log(appreciationPreview);
    try {

      await sp.utility.sendEmail({
        To: [this.state.users.emailId],
        Body:"<div dangerouslySetInnerHTML=__html: "+appreciationPreview.innerHTML+"</div>",
        Subject: "Regarding Appreciation of co-workers",
      });
      alert("Email sent successfully");
    } catch (error) {
      console.log("Error sending email:", error);
    }
    this.setState({initialPage:true,showItem:false, showForm:false})
  }

  public async componentDidMount(){ 
    debugger;
    var reactHandler = this; 
    let Recognition_TitlesArr =[];
    let Mainarr=[];
    this.Recognition_History();
    Recognition_TitlesArr= await this.RetrieveList('Recognition_Titles');
    for(let i=0;i<Recognition_TitlesArr.length;i++){
    //Recognition_TitlesArr.forEach(async function(element){
     
     
      Mainarr.push({
        "Title": Recognition_TitlesArr[i].Title,
         "ID":Recognition_TitlesArr[i].ID,
         "Icon":JSON.parse(Recognition_TitlesArr[i].Icon).serverRelativeUrl
      });
 }
debugger;
    reactHandler.setState({ 
      items: Mainarr
    });
 } 

  public  static onclickofButton(Item){
    debugger;
  
    const body: string = JSON.stringify({
      'Title': Item.Title,
      'Recognition_TitlesId':Item.ID
    });
     
      this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Recognition_History')/items`,
        SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      })
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            response.json().then((responseJSON) => {
              console.log(responseJSON);
              alert(`Item created successfully with ID: ${responseJSON.ID}`);
              
            });
          } else {
            response.json().then((responseJSON) => {
              console.log(responseJSON);
              alert(`Something went wrong! Check the error in the browser console.`);
            });
          }
        }).catch(error => {
          console.log(error);
        });
      
    }
  async RetrieveList(listName) {
    debugger;
    let arr=[];
    await jquery.ajax({ 
      url: `${this.props.siteURL}/_api/web/lists/getbytitle('${listName}')/items`, 
      type: "GET", 
      headers:{'Accept': 'application/json; odata=verbose;'}, 
      success: function(resultData) { 
      
        resultData.d.results.forEach(element => {
          arr.push({
            "Title":element.Title,
            "ID":element.ID,
            "Icon":element.Icon
          })
        });
        
        console.log(resultData); 
      }, 
      error : function(jqXHR, textStatus, errorThrown) { 
      }
  }); 
  return arr;
  }

 
  
  async Recognition_History(){
  debugger;
 
   let arr=[];
   let resultdata;
    await jquery.ajax({ 
        url: `https://integrano.sharepoint.com/sites/NilamDemo/_api/web/lists/getbytitle('Recognition_History')/items?$select=Title,Appreciation_Message,ToId,AuthorId,Recognition_TitlesId`, 
        type: "GET", 
        headers:{'Accept': 'application/json; odata=verbose;'}, 
        success: function(resultData) {
          resultdata=resultData;
        }, 
        error : function(jqXHR, textStatus, errorThrown) { 
        } 
        
    }); 

   // resultdata.d.results.forEach(element => {

      for(let i=0;i<resultdata.d.results.length;i++){
        let element= resultdata.d.results[i];
        let titleData=await this.getTitleInformation(element);
        let TitleInfo=titleData.d.results[0];
        arr.push({
          "Title":element.Title,
          "Appreciation_Message":element.Appreciation_Message,
          "ToId":element.ToId,
          "AuthorId":element.AuthorId,
          "Recognition_TitlesId":element.Recognition_TitlesId,
          "Icon":(TitleInfo.Icon)?JSON.parse(TitleInfo.Icon).serverRelativeUrl:"", 
          "Subject":TitleInfo.Title
        })
      }
    console.log(resultdata,arr); 
    //this.state.history=arr;
   
    this.setState({
      history_1:arr,
      history_1_filter:arr
    })
    return arr; 
  }

   async getTitleInformation(element)
   {
   
      let response;
      await jquery.ajax({ 
        url: `https://integrano.sharepoint.com/sites/NilamDemo/_api/web/lists/getbytitle('Recognition_Titles')/items?$filter=ID eq `+element.Recognition_TitlesId, 
        type: "GET", 
        headers:{'Accept': 'application/json; odata=verbose;'}, 
        success: (resultData,x,x1)=> {
          response=resultData;
          console.log(resultData); 
        
        }, 
        error : function(jqXHR, textStatus, errorThrown) { 
        } 
        
    }); 
    return response;
  }
  static setState(arg0: { items: any; }) {
    throw new Error('Method not implemented.');
  }

  public showRecognitionHistory():void{
    debugger;
    alert("CLick on retrieve all");
    const el = document.getElementById('listid');
    const btn = document.getElementById('Showbutton');

    if (el.style.display === 'none') {
      el.style.display = 'block';
  
      btn.textContent = 'Hide element';
    } else {
      el.style.display = 'none';
  
      btn.textContent = 'Show element';
    }
    
  }

// showcolor =()=>
// {
//  const [items, setItems] = useState<[string, Dispatch<React.SetStateAction<string>>][]>([]);

// }
  
  onchangeofAM=(e)=>{
    debugger;
    this.setState({recognisationItem:{
                                      To:this.state.recognisationItem.To,
                                      Recognition_Titles:this.state.recognisationItem.Recognition_Titles
                                      ,Appreciation_Message:e.target.value.toString(),
                                      Title:this.state.recognisationItem.Title,

                                      From :"",
                                      ImageURL:"",
                                      Subject:""
                                    }});
    this.setsendandPreviewBtnVisibilty(this.state.users.Name,e.target.value,
                                      this.state.recognisationItem.Recognition_Titles);

  }
   
  initialpage = () =>{
    return(
     <section>
     <div>
     <IconButton iconProps={{ iconName: 'Medal' }} title="Medal" ariaLabel="Medal" />Recognition
      <div> <h2>Send recognition to your colleagues</h2></div>
      <div><h5>Show gratitutude over peers who goes above and beyond at work</h5></div>
    
     
      <DefaultButton
           text='Send Recognition'      
           title='Send Recognition'              
          onClick={(e)=>{
               this.setState({showForm:true,initialPage:false,showItem:false});
 
          }}
         /> &nbsp;&nbsp;
     <PrimaryButton //style={{backgroundColor:"blue"}}
           text='Recognition History'      
           title='Recognition History'              
           onClick={(e)=>{
             this.setState({Received:true,Sent:false,showForm:false,showItem:false,initialPage:false});
           }
         }
         />
    </div>    
    </section>
 
 
    )
 
   }

  
  myFun=()=>{
   
   }

   showform = () => {
   
    return (
     
      <div style={{border: '2px solid black'}}>
        <div><IconButton iconProps={{ iconName: 'NavigateBack' }} title="NavigateBack" ariaLabel="NavigateBack" 
        onClick={()=>{
          this.setState({initialPage:true,showForm:false,users:null});
          this.setState({recognitionmsg:null});
          this.state.recognisationItem.Appreciation_Message="";    
        }}/>
        <div className="Header"  style={{ paddingLeft: "25px",paddingRight:"25px"}}><b><strong> Recognition</strong></b>
         </div>
       <div   style={{ paddingLeft: "25px",paddingRight:"25px"}}>
         <PeoplePicker  
      context={this.props.context}  
      titleText="To"  
      personSelectionLimit={3}  
      showtooltip={true}  
      required={true}  
      disabled={false}  
      onChange={this._getPeoplePickerItems}  
      showHiddenInUI={false}  
      ensureUser={true}  
      principalTypes={[PrincipalType.User]}  
      resolveDelay={1000}
      placeholder="Type a name" 
      defaultSelectedUsers={(this.state.users)?([this.state.users.emailId]):null}
      />  
      </div>
      <div className={styles.interactiveWhiteBoard} style={{backgroundImage:"linear-gradient(to right, aliceblue, palevioletred",paddingLeft: "25px",paddingRight:"25px"}}>  
     
      <div id="listid" style={{marginTop:"10px"}}> 
      <label>Title</label>
      <div className={styles.tableStyle} >   
    
      {this.state.items.map(function(item,key){ 
     
      return( 
      
      <div>
        <div className={styles.dropdown} onClick={(test)=>{
                this.setState({recognitionmsg:item.ID,
                  recognisationItem:{
                    To:this.state.recognisationItem.To,
                    Recognition_Titles:item.Icon,
                    Title:item.Title,
                    Appreciation_Message:this.state.recognisationItem.Appreciation_Message
              }});  
              this.setsendandPreviewBtnVisibilty(this.state.users.Name,this.state.recognisationItem.Appreciation_Message,
                item.Icon);
               
              {this.state.items.map(function(item1,key1){ 
                $("#"+item1.ID+"test").css("background-color","");    
              })}
              $("#"+item.ID+"test").css("background-color","red");
              //  this.setState({selectedbuttonId:item.ID.toString()});
             
              //  console.log(this.state.selectedbuttonId);
              
        }}>
         <button className={styles.dropbtn} id={item.ID+"test"}  onClick={()=>
         {
          this.setstate({selectbutton:true})
         }}  style={{backgroundColor:this.state.selectbutton==true?'red':''}}>
          
         <img src={"https://integrano.sharepoint.com/"+item.Icon} className={styles.imageprop}/> &nbsp;&nbsp;
          {item.Title}
         
        </button>
        
         {/* <button className={this.state.clicked ? 'clicked' : ''} onClick={this.setState({clicked:true})} style={{border:"2px solid black"}}></button>
         <button className={styles.dropbtn}/>  */}
      
        
      </div> 
     
      </div>
      
      );
   },this)}
      </div>                  
      </div> 
   

      <div className="ms-TextField ms-TextField--multiline">
              <label className="ms-Label">Message for appreciation (Optional).</label>
              <TextField multiline autoAdjustHeight placeholder='Value here'  id="Appreciation_Message" name="textareaField" onChange={this.onchangeofAM}  defaultValue={this.state.recognisationItem.Appreciation_Message}
              maxLength={100}/>
<span>{this.state.recognisationItem.Appreciation_Message.length}/100</span>
              {/* <textarea className="ms-TextField-field" id="Appreciation_Message" onChange={this.onchangeofAM} value={this.state.recognisationItem.Appreciation_Message}></textarea>
               */}
              
            
     </div>

      {/* { this.state.carouselItems && this.state.carouselItems.length ? 
     
        <Carousel  
          buttonsLocation={CarouselButtonsLocation.center}  
          buttonsDisplay={CarouselButtonsDisplay.buttonsOnly}  
          contentContainerStyles={styles.carouselContent}  
          isInfinite={false}  
          indicatorShape={CarouselIndicatorShape.circle}  
          pauseOnHover={true}  
          element={this.state.carouselItems}  
          
          containerButtonsStyles={styles.carouselButtonsContainer}  
          onRenderIndicator ={this.getselectedAppreciation.bind(this)}
          interval={null}
          onSelect={(index)=>{
            console.log(this.state.carouselItems[index]); 
            console.log(this.state.carouselItems[index].title);  
            this.setState({recognisationItem:{
                  To:this.state.recognisationItem.To,
                  Recognition_Titles:this.state.carouselItems[index].imageSrc,
                  Title:this.state.carouselItems[index].title,
                  Appreciation_Message:this.state.recognisationItem.Appreciation_Message
            }});
            this.setState({ recognitionmsg:this.state.carouselItems[index].id})}}//this.getselectedAppreciation(this.TitleItem)}
         //triggerPageEvent={this.triggerNextElement}
         

        //  onMoveNextClicked={(index: number) => { console.log(`Next button clicked: ${index}`); alert(this.state.carouselItems[index].title) }}
        //  onMovePrevClicked={(index: number) => { console.log(`Prev button clicked: ${index}`); alert(this.state.carouselItems[index].title)}}
         />  
        
        : <p>{this.state.errorMessage}</p>  
      }   */}   </div>

      
     <p className={styles.welcome}>
      <div className="button" style={{float:"right" , paddingLeft: "25px",paddingRight:"100px"}}>
      <DefaultButton
              text='Preview'      
              title='Preview'  
                  
             onClick={(e)=>{
                  this.setState({showForm:false,showItem:true});
             }}
             disabled={this.state.disabledbtns}/> &nbsp;&nbsp;
        <PrimaryButton //style={{backgroundColor:"blue"}}
              text='Send'      
              title='Send'              
             onClick={this.createItem.bind(this)
               || this.setState({showForm:false, initialPage:true})
            }
            disabled={this.state.disabledbtns}/>

          
      
  
   
  </div> 

  </p> </div>  
   </div>
    );
    
  };

  showItem=()=>{
    debugger;
    return(
      <div className="ms-Grid"  style={{border:"2px solid black"}}>

      <div id="recognitionId">
      <div className="Header" style={{ paddingLeft: "25px",paddingRight:"25px"}}><b><strong> Recognition</strong></b></div>
      <div className="ms-Grid-row" style={{backgroundImage:"linear-gradient(to right, aliceblue, palevioletred"}}>
      <div style={{ paddingLeft: "25px",paddingRight:"25px"}}>
          <div id="image">
         {/* <img src="cid:{https://integrano.sharepoint.com//sites/NilamDemo/SiteAssets/Lists/daf14bcf-2ab0-494c-8028-1191d2c65eea/Thanks.png}" width="500" height="500"></img>  */}
         <img src={"https://integrano.sharepoint.com/"+this.state.recognisationItem.Recognition_Titles} className="ImageCls" style={{borderRadius:"50%",border:"2px double grey"}}/>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-row1" style={{fontSize:"20px"}}><b>{this.state.users.Name}</b></div>
            <div className="ms-Grid-row2" style={{fontSize:"20px"}}><b>{this.state.recognisationItem.Title}</b></div>
            <div className="ms-Grid-row3">{this.state.recognisationItem.Appreciation_Message}</div>
            <div className="ms-Grid-row4" style={{marginTop:"20px"}}> From {this.props.context.pageContext.user.displayName}</div>  
          </div>
      </div> 
      </div>
      </div>
        <div style={{ paddingLeft: "25px",paddingRight:"25px",float:"right", marginTop: "10px",}}>
          
        <div className="button" style={{float:"right"}}>
     
      <DefaultButton
              text='Edit'      
              title='Edit'              
             onClick={(e)=>{
              debugger;
                  this.setState({showForm:true,showItem:false});
                  $("#"+this.state.selectedbuttonId+"test").css("background-color","red");
                  console.log(this.state.selectedbuttonId)

             }}
            /> &nbsp;&nbsp;
        <PrimaryButton //style={{backgroundColor:"blue"}}
              text='Send'      
              title='Send'              
             onClick={this.sendEmail.bind(this)}
             onChange={(e)=>{this.setState({initialPage:true,showForm:false, showItem:false})}}
            />

             </div>
          </div>
    </div>
   );
  }

  Sent() {
    return(
      <div>
          <div className="button" style={{float:"right", padding:"right", marginLeft:"60px"}}>
         </div>
        <div><IconButton iconProps={{ iconName: 'NavigateBack' }} title="NavigateBack" ariaLabel="NavigateBack" onClick={()=>{
          this.setState({showItem:true,Received:false})
        }}/>
       
         
      <div className="ms-Grid-row" style={{backgroundColor:"grey",padding: "3px"}}>
      <div style={{ paddingLeft: "25px",paddingRight:"25px",float:"right", marginTop: "10px"}}>
          
     
          </div>
    </div>
    </div>
      <div className="ms-Grid"  style={{display:"flex",width:"20%", height:"20%"}}>
    
          
      
          {this.state.history_1_filter.map(function(item,key){ 
            
           return( <div style={{ paddingLeft: "25px",paddingRight:"25px",height:"450%",width:"470%"}}>
       
           <div className="ms-PersonaCard" style={{backgroundImage:"linear-gradient(to right, aliceblue, palevioletred"}}>
           <div className="ms-PersonaCard-persona">
             <div className="ms-Persona ms-Persona--lg">
               <div className="ms-Persona-imageArea">
                 <div className="ms-Persona-initials ms-Persona-initials--blue"></div>
                 
                 <img className="ms-Persona-image" src={"https://integrano.sharepoint.com/"+item.Icon} style={{backgroundImage:"linear-gradient(to right, aliceblue, palevioletred",borderRadius:"50%",border:"2px double grey"}}/>
                 <div className="ms-Grid-row">
                 <div className="ms-Grid-row1" style={{fontSize:"10px"}}><b>{(item!==null)?item.Title:""}</b></div>
                 <div className="ms-Grid-row2" style={{fontSize:"12px"}}><b><p></p>{(item!==null)?item.Appreciation_Message:""}</b></div>
              
                 </div>
               </div>
             
               </div>
             </div>
           </div>

       </div>
       
       )
           })}
         
      
      </div>
      </div>
     )

  }

showRecords(){
  let userid =this.props.context.pageContext.legacyPageContext.userId;
  return(
    <div>
      <DefaultButton
                      text='Recieved'      
                      title='Recieved'  
                      onClick={()=>{
                        this.setState({history_1_filter:this.state.history_1.filter(function(item){
                          return item.ToId.results[0]==userid;
                          
                        })})
                        
                      }}            
                    /> &nbsp;&nbsp;
              <DefaultButton //style={{backgroundColor:"blue"}}
                    text='sent'      
                    title='sent'            
                    onClick={()=>{
                      debugger;
                      this.setState({history_1_filter:this.state.history_1.filter(function(item){
                        return item.AuthorId==userid;
                      })})
                    }}  
                  /> 
    </div>
  )
}
   Recieved(){
    let userid =this.props.context.pageContext.legacyPageContext.userId;
    
     debugger;  
     return(
      <div>
          <div className="button" style={{float:"right", padding:"right", marginLeft:"60px",borderRadius:"50%"}}>
          <Dropdown 
                            
                placeholder="All Praises"
                label="All Praises"
                options={[
                  { key: 'Recieved', text: 'Recieved'},
                  { key: 'Send', text: 'Send' },
                ]}
               
                onChange={e=>{
                  this.showRecords.bind(this)
                  if(e.target["innerText"].startsWith("Send"))
                  {
                     this.setState({history_1_filter:this.state.history_1.filter(function(item){
                        return item.AuthorId==userid;
                        })})
                  }
                  else {
                      this.setState({history_1_filter:this.state.history_1.filter(function(item){

                      return item.ToId.results[0]==userid;})});
                    }
                
                }   }

                  
                />
              {/* <DefaultButton
                      text='Recieved'      
                      title='Recieved'  
                      onClick={()=>{
                        this.setState({history_1_filter:this.state.history_1.filter(function(item){
                          return item.ToId.results[0]==userid;
                          
                        })})
                      }}            
                    /> &nbsp;&nbsp;
              <DefaultButton //style={{backgroundColor:"blue"}}
                    text='sent'      
                    title='sent'              
                    onClick={()=>{
                      debugger;
                      this.setState({history_1_filter:this.state.history_1.filter(function(item){
                        return item.AuthorId==userid;
                      })})
                    }}  
                  /> */}
          </div>
        <div><IconButton iconProps={{ iconName: 'NavigateBack' }} title="NavigateBack" ariaLabel="NavigateBack" onClick={()=>{
          this.setState({initialPage:true,showItem:false,Received:false})
        }}/>
       
         
      <div className="ms-Grid-row" style={{backgroundColor:"grey",padding: "3px"}}>
      <div style={{ paddingLeft: "25px",paddingRight:"25px",float:"right", marginTop: "10px"}}>
          
     
          </div>
    </div>
    </div>
      <div className="ms-Grid"  style={{display:"flex",width:"20%", height:"20%"}}>
    
          
      
          {this.state.history_1_filter.map(function(item,key){ 
            
           return( <div style={{ paddingLeft: "25px",paddingRight:"25px",height:"450%",width:"470%"}}>
       
           <div className="ms-PersonaCard" style={{backgroundImage:"linear-gradient(to right, aliceblue, palevioletred"}}>
           <div className="ms-PersonaCard-persona">
             <div className="ms-Persona ms-Persona--lg">
               <div className="ms-Persona-imageArea">
                 <div className="ms-Persona-initials ms-Persona-initials--blue"></div>
                 
                 <img className="ms-Persona-image" src={"https://integrano.sharepoint.com/"+item.Icon} style={{backgroundImage:"linear-gradient(to right, aliceblue, palevioletred",borderRadius:"50%",border:"2px double grey"}}/>
                 <div className="ms-Grid-row">
                 <div className="ms-Grid-row1" style={{fontSize:"10px"}}><b>{(item!==null)?item.Title:""}</b></div>
                 <div className="ms-Grid-row2" style={{fontSize:"12px"}}><b><p></p>{(item!==null)?item.Appreciation_Message:""}</b></div>
              
                 </div>
               </div>
             
               </div>
             </div>
           </div>

       </div>
       
       )
           })}
         
      
      </div>
      </div>
     )
 }

  
  private createItem = (): void => {
  
    const body: string = JSON.stringify({
      'Title': this.state.users.Name,

      'ToId': [this.state.users.id],
     'Recognition_TitlesId':this.state.recognitionmsg.toString(),
     
      'Appreciation_Message':document.getElementById("Appreciation_Message")['value']
    });
    this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Recognition_History')/items`,
      SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: body
    })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((responseJSON) => {
            debugger;
            console.log(responseJSON);
           // alert(`Item created successfully with ID: ${responseJSON.ID}`);
            
          });
        } else {
          response.json().then((responseJSON) => {
            console.log(responseJSON);
            alert(`Something went wrong! Check the error in the browser console.`);
          });
        }
      }).catch(error => {
        console.log(error);
      });
      {
        this.showItem();
        this.sendEmail();
      }
      
  }
 
  private _getrecognisationItem(): Promise<ITitleItem[]> {
     debugger;
    let query="/_api/web/lists/getbytitle('Recognition_Titles')/items";

    const URL: string = this.props.siteURL + query;
    return this.props.context.spHttpClient.get(URL,SPHttpClient.configurations.v1)
    .then(response => {
    return response.json();
    })
    .then(json => {
    return json.value;
    }) as Promise<ITitleItem[]>;
    }

    public async getCarouselItems() { 
         debugger;
        // if (this.props.listName) {  

          let query="/_api/web/lists/getbytitle('Recognition_Titles')/items";

          const URL: string = this.props.siteURL + query;
          await this.props.context.spHttpClient.get(URL,SPHttpClient.configurations.v1)
          .then(response => {
            let cardsdata: any[] = [];
            response.json().then((responseJSON)=>{
              console.log(this.getCarouselItems);
              let getAllItems=[];  
              responseJSON.value.map((item:any,index:any)=>{
                getAllItems.push({
                  imageSrc: JSON.parse(item.Icon).serverRelativeUrl,  
                  id:item.Id,
                  title: item.Title,  
                  imageFit: ImageFit.cover  
                })
                cardsdata.push({
                  thumbnail: this.props.siteURL + '/_layouts/15/getpreview.ashx?resolution=1&path=' + encodeURIComponent(JSON.parse(item.Icon).serverRelativeUrl),
                  title: item.Title,
                  location: "SharePoint",
                  url: JSON.parse(item.Icon).serverRelativeUrl
                })
               
              })
              let cardsElements: JSX.Element[] = [];
               
          cardsdata.forEach(item => {
            const previewProps: any = {
              previewImages: [
                {
                  previewImageSrc: item.thumbnail,
                  imageFit: ImageFit.cover,
                  height: 130
                }
              ]
            };
            cardsElements.push(<div>
              <DocumentCard
                type={DocumentCardType.normal}
                onClick={(ev: React.SyntheticEvent<HTMLElement>) => alert("You clicked on a grid item")}>
                <DocumentCardPreview {...previewProps} />
                <DocumentCardDetails>
                  <DocumentCardTitle
                    title={item.title}
                    shouldTruncate={true} />
                </DocumentCardDetails>
              </DocumentCard>
            </div>);
          });
  

              this.setState({carouselJSXItems:cardsElements});
              this.setState({carouselItems:getAllItems});
             
            })
          })
          
          // let carouselItems = await this.SPService.getListItems(this.props.listName);  
          // let carouselItemsMapping = carouselItems.map(e => ({  
          //   imageSrc: JSON.parse(e.Image).serverRelativeUrl,  
          //   title: e.Title,  
          //   description: e.Description,  
          //   showDetailsOnHover: true,  
          //   url: JSON.parse(e.Image).serverRelativeUrl,  
          //   imageFit: ImageFit.cover  
          // }));  
          // this.setState({ listItems: carouselItemsMapping });  
        // }  
        // else {  
        //   this.setState({ errorMessage: "Please set proper list name in property pane configuration." })  
        // }  
      }  
    
  public bindDetailsList(message: string) : void {
    this._getrecognisationItem().then(recognisationItem => {
      this.setState({ TitleItems: recognisationItem,status: message});
    });
    this._getPeoplePickerItems;
  }
 
  

  

  
  public getselectedAppreciation(onClickItem: string){

    console.log(onClickItem);
  }
  public render(): React.ReactElement<IInteractiveWhiteBoardProps>  {
  
    const dropdownRef = React.createRef<IDropdown>();

    return (
      <section>
      <div className={styles.welcome}>
          {this.state.showForm?this.showform():null}
          {this.state.showItem?this.showItem():null}
          {this.state.initialPage?this.initialpage():null}
          {(this.state.Received)?this.Recieved():null}
          
          {/* {
          (this.state.selectedbuttonId!=="") &&
            $("#"+this.state.selectedbuttonId+"test").css("background-color","red")
          } */}
     </div>
   </section>
 );
}



      
   
   
   @autobind  
   private _getPeoplePickerItems(items: any[]) {  
    debugger; 
   
     for (let item in items) {  
          this.setState({ users: {
            id:items[item].id,
            Name:items[item].text,
            emailId:items[item].secondaryText,
            resultdata:items[item].resultData
          } });  
     }  
     this.state.recognisationItem.Title=this.state.users.Name;
     this.setsendandPreviewBtnVisibilty(items[0],this.state.recognisationItem.Appreciation_Message,
      this.state.recognisationItem.Recognition_Titles);
     console.log(this.state.recognisationItem.Title);
    
   }
   
   private  setsendandPreviewBtnVisibilty=(userName,appreciationMsg,Recognition_Titles)=>{

    if(userName && appreciationMsg && Recognition_Titles ){
             this.setState({disabledbtns:false});
      }
      else{
        this.setState({disabledbtns:true});
      }
     }

   }

 
    
    
  
  
  
 
 

 

  