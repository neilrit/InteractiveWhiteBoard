import { IColumn, MessageBarType } from "office-ui-fabric-react";
import { IRecognitionHistoryItem } from "./IRecognitionHistoryItem";
import { ITitleItem } from "./ITitleItem";

export interface IInteractiveWhiteBoardState{
    
    status:string;
    IconItem:[];
    items:any[];
    buttonSet:{
        Title:string,
        Icon:string
    },

    history_1:any[],
    columns: IColumn[];
    TitleItems: ITitleItem[];
    TitleItem:any[];
    RecognisationItems:IRecognitionHistoryItem[];
    recognisationItem:IRecognitionHistoryItem;
    errorMessage: string;  
    title: string;  
    users: {
        id:string,
        Name:string,
        emailId:string,
        resultdata:[],
    };  
    showMessageBar: boolean;  
    messageType?: MessageBarType;  
    message?: string; 
    carouselItems:any[]
    carouselJSXItems:JSX.Element[];
    recognitionmsg:string,
    //resultData:[];
    clicked: boolean,


    isEditFormOpen: false,
    showForm:boolean,
    showItem:boolean,
    Received:boolean,
    Sent:boolean,
    initialPage:boolean,
    history_1_filter:any[],
    disabledbtns:boolean,
    selectedbuttonId:string,
    selectbutton:boolean
  
}