import * as React from 'react';

import styles from './DeskBookingTool.module.scss';
import { IDeskBookingToolProps } from './IDeskBookingToolProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Service from './Service'
import { Item, Items } from '@pnp/sp/items';
import { sp, toAbsoluteUrl } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";

//import DatePicker from "react-datepicker";

import "react-datepicker/dist/react-datepicker.css";

import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { Grid, Paper, Table, ModalManager } from '@material-ui/core';

import {
  TextField,ICheckboxStyles,IChoiceGroupOption,ChoiceGroup,
  Stack, IDropdownOption, Dropdown, IDropdownStyles,Link,IconButton,Checkbox,
  IStackStyles, DatePicker, PrimaryButton, Label, getHighContrastNoAdjustStyle, IStackTokens, StackItem
} from '@fluentui/react';




import { DateTimePickerComponent } from '@syncfusion/ej2-react-calendars';


const sectionStackTokens: IStackTokens = { childrenGap: 10 };
const sectionStackTokens1: IStackTokens = { childrenGap: 5 };
const stackTokens = { childrenGap: 80 };
const stackStyles: Partial<IStackStyles> = { root: { padding: 10 } };
const stackButtonStyles: Partial<IStackStyles> = { root: { Width: 20 } };

const RadioExp: IChoiceGroupOption[] = 

[  { key: "Yes", text: "Yes" , },  { key: "No", text: "No" },];  

const Radio14Days: IChoiceGroupOption[] = 

[  { key: "Yes", text: "Yes" , },  { key: "No", text: "No" },];  

const RadioClose: IChoiceGroupOption[] = 

[  { key: "Yes", text: "Yes" , },  { key: "No", text: "No" },];  

const RadioTestPos: IChoiceGroupOption[] = 

[  { key: "Yes", text: "Yes" , },  { key: "No", text: "No" },];  


let GeneralDeskLists = [];
let FixedDeskLists = []
let TotalSeatsList = [];
let SelLoca = '';
let SelBuilding = '';
let SelBookingType = ''
let SelFloorLevel = ''
let BlockCount = null;
let FinalTotalSeats = [];
let BookedSeatsList = [];
let ConcString = '';
let RootUrl = '';
let Userreqemail = '';
let UserreqName = '';
let Markup = '';
let UserNameCollection=[];
let AllCheckedItems: any = [];
var AllDelrecordIds: any = [];

//const [selectedDate, handleDateChange] = useState(new Date());


const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
  

};



let checkboxstyles: ICheckboxStyles = { 
  root: { 
    marginTop: '10px',  
    
    paddingTop: '10px',  
    paddingBottom: '10px',  
    paddingLeft: '10px'  
  }, 
  checkmark: { 

    backgroundColor:'green'
  } 
}; 




const options: IDropdownOption[] = [

    {key:'Select',text:'Select'},
    { key: '1', text: '1' },
    { key: '2', text: '2' },
    { key: '3', text: '3' },
    { key: '4', text: '4' },
    { key: '5', text: '5' },
    { key: '6', text: '6' },

  ];

 
  



export interface IDeskBookingFieldsState {
  Mylocationval: any;
  MyBuildingVal: any;
  MyBookingTypeVal: any;
  MyFloorType: any;
  flag: boolean;
  procflag:boolean;
  StartDate: any;
  EndDate: any;
  BlockCount: number;
  ItemInfo: any;
  ConcString: string;
  LocationListItems: any;
  BuildingListItems: any;
  BookingListItems: any;
  FloorListItems: any;
  userExsits: boolean;
  userEmail: any;
  deskId: any;
  FinalMarkup: any;
  GridUserValues:any;
  chckboxesseats:any;
  Dispalygrid:boolean;
  checkedArray:any;
  MyUserName:any;
  uncheckedArray:any;
  checkstatus:boolean;
  DelRecIds:any;
  NoofSeats:any;
  UserLoginName:any;
  MaximumDate:any;
  MinDate:any;
  MyStrUser:any;
  UserDefValue:any;
  LinkFloor:any;
  DeskDesc:any;
  SelDeskreqsId:any;
  DefalSelcarray:any;
  Userboolval:boolean;
  Exp:any;
  MyDays:any;
  Closekey:any;
  QuestKey:any;
  isDisable:boolean;
  FirstDivVisble:boolean;

}

export default class DeskBookingTool extends React.Component<IDeskBookingToolProps, IDeskBookingFieldsState> {

  public _service: any;
  public GlobalService: any;

  protected ppl;

  public constructor(props: IDeskBookingToolProps) {
  super(props);
  
    this.state = {

      Mylocationval: null,
      MyBuildingVal: null,
      MyBookingTypeVal: null,
      MyFloorType: null,
      flag: false,
      procflag:false,
      NoofSeats: null,
      StartDate: null,
      EndDate: null,
      BlockCount: null,
      ItemInfo: null,
      ConcString: "",
      LocationListItems: [],
      BuildingListItems: [],
      BookingListItems: [],
      FloorListItems: [],
      userExsits: false,
      userEmail: "",
      deskId: "",
      FinalMarkup: "",
      GridUserValues:[],
     chckboxesseats:[],
     Dispalygrid:false,
     checkedArray:[],
     MyUserName:"",
     uncheckedArray:[],
     checkstatus:false,
     DelRecIds:[],
     UserLoginName:"",
     MaximumDate:null,
     MinDate:"",
     MyStrUser:[],
     UserDefValue:[],
     LinkFloor:"",
     DeskDesc:"",
     SelDeskreqsId:"",
     DefalSelcarray:[],
     Userboolval:false,
     Exp:"",
     MyDays:"",
     Closekey:"",
     QuestKey:"",
     isDisable:true,
     FirstDivVisble:true
     

    };


    

    RootUrl = this.props.url;

    this._service = new Service(this.props.url, this.props.context);

    this.GlobalService = new Service(this.props.url, this.props.context);

    this.getAllLocations();

    this.getUserDetails();

    alert('five');
    
    

  }

  private ChangeSeats(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption):void{
    this.setState({NoofSeats:item});
  }

  private GetStartDateandEndDate()
  {

    

    let StartAM='9:00 AM'
    let EndPm='5:00 PM'
    let now = new Date();
    now.setDate(now.getDate()+1)
    let day = ("0" + now.getDate()).slice(-2);
    let month = ("0" + (now.getMonth() + 1)).slice(-2);
    let today = (month) + "/" + (day) + "/" + now.getFullYear()+" "+" "+StartAM;
    let today1 = (month) + "/" + (day) + "/" + now.getFullYear()+" "+" "+EndPm;

    

    
    let now1 = new Date();
    now1.setDate(now1.getDate()+30)

    let day1 = ("0" + now1.getDate()).slice(-2);

    let month1=("0" + (now1.getMonth() + 1)).slice(-2);

   let todaymax=(month1) + "/" + (day1) + "/" + now1.getFullYear()+" "+" "+StartAM;

  
   

    

    let now2=new Date();
    now2.setDate(now2.getDate())

    let day2 = (now2.getDate());

    let month2=("0" + (now2.getMonth() + 1)).slice(-2);

    let todaymin=(month2) + "/" + (day2) + "/" + now2.getFullYear()+" "+" "+StartAM;


  
    

   this.setState({ StartDate: today });

   this.setState({ EndDate: today1 });

   this.setState({MaximumDate:todaymax});

   this.setState({MinDate:todaymin});



  }


  public async getAllLocations() {

    var myLocationLocal: any = [];

    var data = await this._service.MyGetAllocations();

    console.log(data);

    var AllLocations: any = [];

    for (var k in data) {

      AllLocations.push({ key: data[k].Title, text: data[k].Title });
    }

    console.log(AllLocations);

    AllLocations.map(item => {
      let Itemexsits = false;
      if (myLocationLocal != null) {
        if (myLocationLocal && myLocationLocal.length > 0) {
          myLocationLocal.map(ditem => {
            if (ditem.key === item.key) {

              Itemexsits = true;
            }

          });
        }

      }

      if (!Itemexsits) {


        myLocationLocal.push({ key: item.key, text: item.text });
      }


    });
   this.setState({ LocationListItems: myLocationLocal });

  }


  private async GetBulidingData(SelLocVal: string) {

    var myBuildingLocal: any = [];

    var data = await this._service.MyGetBulidingData(SelLocVal);

    var AllBuildings: any = [];

    let BulingLevel = data[0].BulidingName;

    let arr = BulingLevel.split(',')

    for (var k in arr) {
      AllBuildings.push({ key: arr[k], text: arr[k] });
    }

    console.log(AllBuildings);

    AllBuildings.map(item => {
      let Itemexsits = false;

      if (myBuildingLocal != null) {
        if (myBuildingLocal && myBuildingLocal.length > 0) {

          myBuildingLocal.map(ditem => {
            if (ditem.key === item.key) {

              Itemexsits = true;
            }

          });
        }

        if (!Itemexsits) {

          myBuildingLocal.push({ key: item.key, text: item.text });
        }

      }
    });


    this.setState({ BuildingListItems: myBuildingLocal });

  }

  //Start



  private async GetBookingsData(LocaVal: string, BuildNamVal: string) {

    
    this.GlobalService = new Service(this.props.url, this.props.context);

    let MyUrl = await this.GlobalService.GetUrls(this.state.Mylocationval);

    //alert(MyUrl);

    this._service = new Service(MyUrl, this.props.context);

    let mycurgroup = await this._service.getCurrentUserSiteGroups();


    for (let grpcount = 0; grpcount < mycurgroup.length; grpcount++) {

      if(this.state.Mylocationval=='Canada')
      {

        if (mycurgroup[grpcount].Title == 'FixedDeskUsersCanada')
        {
          this.setState({ userExsits: true });

        }

      }

      if(this.state.Mylocationval=='Brazil')
      {

        if (mycurgroup[grpcount].Title == 'FixedDeskUsersBrazil')
        {
          this.setState({ userExsits: true });

        }

      }

      if(this.state.Mylocationval=='Orlando')
      {

        if (mycurgroup[grpcount].Title == 'FixedDeskUsersOrlando')
        {
          this.setState({ userExsits: true });

        }

      }

      if(this.state.Mylocationval=='NewYork')
      {

        if (mycurgroup[grpcount].Title == 'FixedDeskUsersNewYork')
        {
          this.setState({ userExsits: true });

        }

      }

      if(this.state.Mylocationval=='Charlotte')
      {

        if (mycurgroup[grpcount].Title == 'FixedDeskUsersCharlotte')
        {
          this.setState({ userExsits: true });

        }

      }






   //END
    }


    var myBookingLocal: any = [];


    this._service = new Service(RootUrl, this.props.context);

    var data = await this._service.MyGetBookingType(LocaVal, BuildNamVal);

    let BookingLevel;

    let arr :any=[];

    var AllBookings: any = [];

    for( var count in  data)
    {

     BookingLevel = data[count].BookingType;

     arr = BookingLevel.split(',')

    }
    

    for (var k in arr) {

      if(this.state.userExsits==true)
      {

      AllBookings.push({ key: arr[k], text: arr[k] });
        

      }

      if(this.state.userExsits==false)
      {

        if(arr[k]!='Fixed Desk')
        {


       AllBookings.push({ key: arr[k], text: arr[k] });
        }

      }


    }


    AllBookings.map(item => {
      let Itemexsits = false;

      if (myBookingLocal != null) {
        if (myBookingLocal && myBookingLocal.length > 0) {
          myBookingLocal.map(ditem => {
            if (ditem.key === item.key) {

              Itemexsits = true;
            }

          });
        }

        if (!Itemexsits) {

       myBookingLocal.push({ key: item.key, text: item.text });

       }

      }
    });

    this.setState({ BookingListItems: myBookingLocal });

  }


  //END

  
  private async GetFloorsData(LocaVal: string, BuildNamVal: string, BookingTypeval: string) {

    var myFloorLocal: any = [];

    var data = await this._service.MyGetFloorsData(LocaVal, BuildNamVal, BookingTypeval);

    var AllFloors: any = [];

    let FloorLevel = data[0].FloorLevel;

    console.log(FloorLevel);

    let arr = FloorLevel.split(',')

    console.log(arr[0]);

    for (var k in arr) {

      AllFloors.push({ key: arr[k], text: arr[k] });

    }

    AllFloors.map(item => {
      let Itemexsits = false;

      if (myFloorLocal != null) {
        if (myFloorLocal && myFloorLocal.length > 0) {
          this.state.BookingListItems.map(ditem => {
            if (ditem.key === item.key) {

              Itemexsits = true;
            }

          });
        }

        if (!Itemexsits) {

        myFloorLocal.push({ key: item.key, text: item.text });
        }

      }
    });

    this.setState({ FloorListItems: myFloorLocal });


  }

 


  private handleChangeLocation(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {

    if (this.state.BuildingListItems.length > 0) {

      BuildingListItems: [];
      BookingListItems: [];
      FloorListItems: [];

      this.setState({ BuildingListItems: [] });
      this.setState({ BookingListItems: [] });
      this.setState({ FloorListItems: [] });
      this.setState({ ConcString: '' });

      this.setState({ MyBuildingVal: 'Select' });
      this.setState({ MyBookingTypeVal: 'Select' });
      this.setState({ MyFloorType: 'Select' });


    }


    this.GetBulidingData(item.text);

    SelLoca = item.text;

    this.setState({ Mylocationval: item.text });

    console.log(this.state.MyBuildingVal);

  }

  private handleChangeBuilding(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {


    if (this.state.BookingListItems.length > 0) {

       BookingListItems: [];
      FloorListItems: [];
      this.setState({ ConcString: '' });
      this.setState({ BookingListItems: [] });
      this.setState({ FloorListItems: [] });
      this.setState({ MyBookingTypeVal: 'Select' });
      this.setState({ MyFloorType: 'Select' });

    }



    console.log(item.text);
    SelBuilding = item.text;

    this.setState({ MyBuildingVal: item.text });
    this.GetBookingsData(SelLoca, SelBuilding);


  }

  private handleChangeBookingType(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {


    if (this.state.FloorListItems.length > 0) {

       FloorListItems: [];
      this.setState({ FloorListItems: [] });
      this.setState({ MyFloorType: 'Select' });
      this.setState({ ConcString: '' });

    }


    if (this.state.FloorListItems.length > 0) {
      FloorListItems: [];

    }

    console.log(item.text);
    SelBookingType = item.text;
    this.setState({ MyBookingTypeVal: item.text });
    this.GetFloorsData(SelLoca, SelBuilding, SelBookingType);

  }


  private handleChangeFloorLevel(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {

    this.setState({ ConcString: '' });
    console.log(item.text);
    SelFloorLevel = item.text;
    this.setState({ MyFloorType: item.text });


  }


  private OnBtnClick(): void {

    if (this.state.Mylocationval == null || this.state.Mylocationval == 'Select Location' || this.state.Mylocationval == 'Select') {

      alert('please select location');
      this.setState({ flag: false });

    }

    else if (this.state.MyBuildingVal == null || this.state.MyBuildingVal == 'Select Building' || this.state.MyBuildingVal == 'Select') {

      alert('please select Building');
      this.setState({ flag: false });

    }


    else if (this.state.MyBookingTypeVal == null || this.state.MyBookingTypeVal == 'Select BookingType' || this.state.MyBookingTypeVal == 'Select') {

      alert('please select BookingType');
      this.setState({ flag: false });

    }


    else if (this.state.MyFloorType == null || this.state.MyFloorType == 'Select  FloorLevel' || this.state.MyFloorType == 'Select') {

      alert('please select FloorLevel');
      this.setState({ flag: false });

    }

    else {

      this.setState({ flag: true });

      this.GetStartDateandEndDate();

    }

  }



  public async SelectSeats() {


    let MyStartDate = this.formatdate(this.state.StartDate);

    let MyEndDate = this.formatdate(this.state.EndDate);

    let MyFloorLevel = this.state.MyFloorType;

    let MyBookingType = this.state.MyBookingTypeVal;

    if (this.state.NoofSeats == null || this.state.NoofSeats == '' || this.state.NoofSeats=='Select') {

      alert('please select No of Seats');

    }


    else if (this.state.StartDate == null || this.state.StartDate == '') {

      alert('Please select Start Date');
    }

    else if (this.state.EndDate == null || this.state.EndDate == '') {

      alert('Please select End Date');

    }

    else if(Date.parse(MyStartDate) > Date.parse(MyEndDate))
    {

        alert('Start Date should be less than End Date');

    }


    else {

      this.GlobalService = new Service(this.props.url, this.props.context);


      let MyUrl = await this.GlobalService.GetUrls(this.state.Mylocationval);

      alert(MyUrl);

      this._service = new Service(MyUrl, this.props.context);

      let BlockStatus = await this._service.CheckBlockDate(MyStartDate, MyEndDate);


      if (BlockStatus == true) {


        alert("Date has been blocked");

      }

      if (BlockStatus == false) {


        var MyBuild=this.state.MyBuildingVal;

        let  MyHyperlink= await this._service.GetPDFLinks1(MyBuild,MyBookingType,MyFloorLevel);

        

        this.setState({LinkFloor:MyHyperlink});

        

        TotalSeatsList = await this._service.TotalNoofSeats(MyFloorLevel, MyBookingType);

        BookedSeatsList = await this._service.BookedSeats(MyFloorLevel, MyBookingType, MyStartDate, MyEndDate);

        if(BookedSeatsList!=null)
        {
         
          var strhtml = '';
          let avaiblesets:any=[];
          for (var seatcount = 0; seatcount < TotalSeatsList.length; seatcount++) {
            let IsItemExist = this.checkItemInArray(TotalSeatsList[seatcount], BookedSeatsList);
            let seatDetails:{};
            if (IsItemExist) {
              seatDetails={seatId:TotalSeatsList[seatcount],seatTaken:true};
            }else{
              seatDetails={seatId:TotalSeatsList[seatcount],seatTaken:false};
            }
            avaiblesets.push(seatDetails);
           
          }
          this.setState({chckboxesseats:avaiblesets});
          this.setState({Dispalygrid:true});

        }

        else
        {

          var strhtml = '';
          let avaiblesets:any=[];
           BookedSeatsList = [];
          for (var seatcount = 0; seatcount < TotalSeatsList.length; seatcount++) {
            let IsItemExist = this.checkItemInArray(TotalSeatsList[seatcount], BookedSeatsList);
            let seatDetails:{};
            if (IsItemExist) {
              seatDetails={seatId:TotalSeatsList[seatcount],seatTaken:true};
            }else{
              seatDetails={seatId:TotalSeatsList[seatcount],seatTaken:false};
            }
            avaiblesets.push(seatDetails);
           
          }
          this.setState({chckboxesseats:avaiblesets});
          this.setState({Dispalygrid:true});

        }

      }

    }

  }


  private async _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);

    if(items.length>0)
    {

    Userreqemail = items[0].secondaryText;
    UserreqName = items[0].text

    }

    //this.ppl.onChange([]);

  }


  public checkItemInArray(itemName: string, ItemArray: any): boolean {
    let isItemExist = false;

    for (var bkseatcount = 0; bkseatcount < ItemArray.length; bkseatcount++) {
      if (ItemArray[bkseatcount] === itemName)
        isItemExist = true;
    }
    return isItemExist;
  }


  public checkItemChecked(itemName: string, ItemArray: any): boolean {
    let isItemchecked = false;

    for (var Count = 0; Count < ItemArray.length; Count++) {
      if (ItemArray[Count].key === itemName)
      isItemchecked = true;
    }
    return isItemchecked;
  }

  public checkItemUserName(itemName: string, ItemArray: any): boolean {
    let isItemchecked = false;

    for (var Count = 0; Count < ItemArray.length; Count++) {
      if (ItemArray[Count].UserEmail === itemName)
      isItemchecked = true;
    }
    return isItemchecked;
  }


  public checkDeskIDchked(itemName: string, ItemArray: any): boolean {
    let isItemchecked = false;

    for (var Count = 0; Count < ItemArray.length; Count++) {
      if (ItemArray[Count].DeskId === itemName)
      isItemchecked = true;
    }
    return isItemchecked;
  }




  private OnBtnPrevClick(): void {

    this.setState({ flag: false });

    this.setState({Dispalygrid:false});

    this.setState({GridUserValues:[]});

    AllCheckedItems=[];



  }

  private OnAddRowsClick(): boolean {

    
    let ISUserNameExsists;
    
    if(this.state.deskId == null || this.state.deskId == '')
    {

       alert('Please enter the DeskId');
       return false;


    }

    else if(Userreqemail==null || Userreqemail=="")
    {

      alert('Please select username');
      return false;
    }

    
   
   else  if(this.state.deskId != null || this.state.deskId != '' )
    {
      let Isitemchecked=this.checkItemChecked(this.state.deskId,AllCheckedItems);
      
      let ISUserNameExsists=this.checkItemUserName(Userreqemail,this.state.GridUserValues);

      let IsDeskIDexsists=this.checkDeskIDchked(this.state.deskId,this.state.GridUserValues);
      

      if(Isitemchecked===false)
      {

        alert('Please enter desk no which you selected above');
        return false;
      }

      if(ISUserNameExsists===true)
      {

        alert('Please Select different userName');
        return false;

      }

      if(IsDeskIDexsists===true)
      {

        alert('Please enter diifernt deskID');
        return false;

      }

    

      if(Isitemchecked===true && ISUserNameExsists===false && IsDeskIDexsists===false)
      {



     console.log(UserNameCollection);

    let GridLocal:any=(this.state.GridUserValues?this.state.GridUserValues:[]);
   
    let GridItem:{}={UserEmail:Userreqemail,UserName:UserreqName,DeskId:this.state.deskId};

    GridLocal.push(GridItem);

    this.setState({GridUserValues:GridLocal});

    this.setState({deskId:''});

    this.setState({Userboolval:true});

    Userreqemail='';

    UserreqName='';

    this.ppl.onChange([]);

    this.setState({MyUserName:''});
   
    this.setState({checkstatus:false});



    this.clear();

    

      }

    }

  }
    
private clear():void{

    AllCheckedItems:[];
    Userreqemail="";

    var MyEmail: any = [];

    this.setState({UserDefValue:MyEmail});

  }


  public handlestartDateChange = (date: any) => {

   

     this.setState({ StartDate: date.value });
   }


   public handleEndDateChange = (date: any) => {

     this.setState({ EndDate: date.value });
   }

  private changeDesk(data: any): void {

    this.setState({ deskId: data.target.value });

  }

  private SelectDelDesIDS(data: any): void {

    
   
   AllDelrecordIds.push({ key: data.target.value, text: data.target.value });

   this.setState({DelRecIds:AllDelrecordIds});

  }

  

  
  private onDeleteClick(Item: any): void {

    //this.setState({ deskId: data.target.value });

    var MyGlobarLocal: any = this.state.GridUserValues; 

    if(MyGlobarLocal.length==1)
    {


      if(MyGlobarLocal[0].DeskId==Item.DeskId)
      {

       

        MyGlobarLocal=[];

     }

    }

    for(var count=0;count<MyGlobarLocal.length;count++)
    {


      if(MyGlobarLocal[count].DeskId==Item.DeskId)
      {

        let Index=count;

        MyGlobarLocal.splice(Index,count);

     }
       

    }

    

    this.setState({GridUserValues:MyGlobarLocal});

    

  }

  private async onSubmitClick()
  {

    
    this.GlobalService = new Service(this.props.url, this.props.context);

    let MyUrl1 = await this.GlobalService.GetUrls(this.state.Mylocationval);

    this._service = new Service(MyUrl1, this.props.context);

    
    for(let count=0;count<this.state.GridUserValues.length;count++)
    {

      let MyTilte=this.state.Mylocationval + "-" +this.state.GridUserValues[count].UserName + "-" +this.state.MyFloorType 

      this._service.onDrop(this.state.Mylocationval,this.state.MyBuildingVal,this.state.MyBookingTypeVal,this.state.MyFloorType,this.state.StartDate,this.state.EndDate,this.state.GridUserValues[count].DeskId,this.state.GridUserValues[count].UserEmail,MyTilte).then(function (data)
      {

     
     window.location.replace("https://capcoinc.sharepoint.com/sites/Global-Capco-Desk-Reservations/");

      });

     

    }

    alert('Submitted Successfully');

    
  }


  private async getUserDetails()
  {
    let result= await this._service.getCurrentUser();

    Userreqemail = result.Email;
    UserreqName = result.Title;

   this.setState({UserLoginName:result.Email});

  

   

  }

  


  private async changechbox(data: any) {


     //New Lines

     console.log(data.target.value);

     let SelDeskId=data.target.attributes["aria-label"].value;

     let Checkstatusval=data.target.checked;

     this.GlobalService = new Service(this.props.url, this.props.context);

    let MyUrl2 = await this.GlobalService.GetUrls(this.state.Mylocationval);

    this._service = new Service(MyUrl2, this.props.context);

    
    let DeskDescprtion= await this._service.GetDeskDesc(SelDeskId);
    

    console.log(DeskDescprtion);

    if(DeskDescprtion!=null)
    {
    this.setState({DeskDesc:DeskDescprtion});

    }

    //END

    

   
    if(Checkstatusval==true)
    {

     
      AllCheckedItems.push({ key: SelDeskId, text: SelDeskId });

    
  }

  if(Checkstatusval==false)
  {

    
  //for(var count=0;count<AllCheckedItems.length;count++)
	//{

    AllCheckedItems.splice(AllCheckedItems.indexOf(SelDeskId), 1);
  //}

  }

 



  }

  public formatdate(strDate) {
    var dt = new Date(strDate);
    return dt.toISOString();
  }


//Radio Events


public ChangeExp=async(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): Promise<void>=> {  

  this.setState({  

    Exp: option.key  

    });  

    if(this.state.Exp=='No' && this.state.MyDays=='No' && this.state.Closekey=='No' && this.state.QuestKey=='No')
    {


      this.setState({  

        isDisable: false
    
        });  

    }

    else
    {
      this.setState({  

        isDisable: true
    
        }); 

    }

  }

  public Change14days=async(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): Promise<void>=> {  

    this.setState({  
  
      MyDays: option.key  
  
      });  

      if(this.state.Exp=='No' && this.state.MyDays=='No' && this.state.Closekey=='No' && this.state.QuestKey=='No')
      {
  
        this.setState({  

          isDisable: false
      
          });  
      }

      else
      {
        this.setState({  

          isDisable: true
      
          }); 

      }
  
    }

    public ChangeClose=async(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): Promise<void>=> {  

      this.setState({  
    
        Closekey: option.key  
    
        });  

        if(this.state.Exp=='No' && this.state.MyDays=='No' && this.state.Closekey=='No' && this.state.QuestKey=='No')
        {
    
          this.setState({  

            isDisable: false
        
            });  
        }

        else
        {
          this.setState({  

            isDisable: true
        
            }); 

        }
    
      }

      public ChangeQuestions=async(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): Promise<void>=> {  

        this.setState({  
      
          QuestKey: option.key  
      
          });  

          if(this.state.Exp=='No' && this.state.MyDays=='No' && this.state.Closekey=='No')
          {
      
            this.setState({  

              isDisable: false
          
              });  
      
          }


          else
          {
            this.setState({  

              isDisable: true
          
              }); 

          }

      
        }


        private OnBtnProcClick() :void {

          this.setState({  
      
            procflag:true,
            FirstDivVisble:false
        
            });  

        }

        private OnBtnCancelClick() :void {


          window.location.replace("https://capcoinc.sharepoint.com/sites/Global-Capco-Desk-Reservation");
        }

//End



  public render(): React.ReactElement<IDeskBookingToolProps> {


    return (


      <Stack tokens={stackTokens} styles={stackStyles}>
      {this.state.flag == false && this.state.FirstDivVisble==true &&
       <Stack>
      <label className={styles.headings}>1. Are you experiencing any of the below symptoms?</label>
      <br></br>
      <label className={styles.headings}>a. Fever and/or chills</label>
      <br></br>
      <label className={styles.headings}>b. Cough</label>
      <br></br>
      <label className={styles.headings}>c. Shortness of breath</label>
      <br></br>
      <label className={styles.headings}>d. Decrease or loss of taste/ smell</label>
      <br></br>
      <label className={styles.headings}>e. Muscle aches/ joint pain</label>
      <br></br>
      <label className={styles.headings}>f. Extreme tiredness</label>
      <br></br>
       <ChoiceGroup className={styles.onlyFont} options={RadioExp} onChange={this.ChangeExp}/>
       <br></br>
       <label className={styles.headings}>2. In the last 14 days, have you travelled outside of Canada and been told to quarantine (per the federal quarantine requirements)?</label>
       <br></br>
       <ChoiceGroup className={styles.onlyFont} options={Radio14Days} onChange={this.Change14days}/>
       <br></br>
       <label className={styles.headings}>3. Have you been in close contact with someone who is sick in the past 14 days?</label>
       <br></br>
       <ChoiceGroup className={styles.onlyFont} options={RadioClose}  onChange={this.ChangeClose}/>
       <br></br>
       <label className={styles.headings}>4. Have you tested positive for COVID-19 in the past 14 days?</label>
       <br></br>
       <ChoiceGroup className={styles.onlyFont} options={RadioTestPos} onChange={this.ChangeQuestions}/>
       <br></br>
       <label className={styles.headings}>If you have answered YES to any of these questions, please DO NOT ENTER the office. Please stay at home and self-isolate. Contact Telehealth or your health provider to find out how to proceed.</label>
       <br></br>
       <Stack horizontal tokens={sectionStackTokens}>
        <StackItem>
        <PrimaryButton disabled={this.state.isDisable} text="Proceed" styles={stackButtonStyles} className={styles.ProceedButton} onClick={this.OnBtnProcClick.bind(this)}/>
       </StackItem>
       <StackItem>
       <PrimaryButton text="Cancel" onClick={this.OnBtnCancelClick.bind(this)} styles={stackButtonStyles} className={styles.button} />
       </StackItem>
       </Stack>
       

  
      </Stack>

      }


        {this.state.flag == false && this.state.procflag==true &&
          <Stack>

            {/* StartAM */}

            {/* <DatePicker

                    onChange={this.handlestartDateChange}

                    selected={this.state.StartDate}

                    showTimeSelect

                    dateFormat="MM/dd/yyyy   HH:mm a"

                /> */}


            {/* END */}
    



            {/* <b><div className={styles.headings}>Location</div></b> */}
            <label className={styles.headings}>Location</label>
            <br></br>
            <Dropdown placeHolder="Select Location" options={this.state.LocationListItems} className={styles.headings} styles={dropdownStyles} selectedKey={this.state.Mylocationval ? this.state.Mylocationval : undefined} onChange={this.handleChangeLocation.bind(this)} /><br></br>
            <label className={styles.headings}>Building</label><br></br>
            <Dropdown placeHolder="Select Building" options={this.state.BuildingListItems} styles={dropdownStyles} selectedKey={this.state.MyBuildingVal ? this.state.MyBuildingVal : undefined} onChange={this.handleChangeBuilding.bind(this)} /><br></br>
            <label className={styles.headings}>Booking Type</label><br></br>
            <Dropdown placeHolder="Select Booking Type" options={this.state.BookingListItems} styles={dropdownStyles} selectedKey={this.state.MyBookingTypeVal ? this.state.MyBookingTypeVal : undefined} onChange={this.handleChangeBookingType.bind(this)} /><br></br>
            <label className={styles.headings}>Floor Level</label><br></br>
            <Dropdown placeHolder="Select Floor Level" options={this.state.FloorListItems} styles={dropdownStyles} selectedKey={this.state.MyFloorType ? this.state.MyFloorType : undefined} onChange={this.handleChangeFloorLevel.bind(this)} /><br></br>
            <PrimaryButton text="Next" onClick={this.OnBtnClick.bind(this)} styles={stackButtonStyles} className={styles.button} />
            </Stack>

        }


        {this.state.flag == true && this.state.procflag &&
          <Stack className={styles.dateTimeClass}>

            <div>
            <label className={styles.headings}>Start Date</label><br></br>
            </div><br></br>
            <DateTimePickerComponent onChange={this.handlestartDateChange} value={this.state.StartDate} max={this.state.MaximumDate} min={this.state.MinDate} ></DateTimePickerComponent>  
            <div>
            <label className={styles.headings}>End Date</label><br></br>
            </div><br></br>
            <div>
            
          <DateTimePickerComponent onChange={this.handleEndDateChange} value={this.state.EndDate} max={this.state.MaximumDate} min={this.state.MinDate}></DateTimePickerComponent> 
          <br></br>
          </div>
              
          <br></br>

            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem className={styles.commonstyle}>

              <label className={styles.headings}>Number of Seats:</label><br></br>
               
              </StackItem>
              <StackItem className={styles.commonstyle}>
              <Dropdown
                placeholder="Select No of seats"
                options={options}
                styles={dropdownStyles}
                selectedKey={this.state.NoofSeats ? this.state.NoofSeats.key : undefined}
                onChange={this.ChangeSeats.bind(this)} 
            />   

            </StackItem>
            </Stack>

            <br></br>

            <Stack horizontal tokens={sectionStackTokens}>
            <StackItem className={styles.commonstyle}>
            <PrimaryButton text="Previous" styles={stackButtonStyles}  onClick={this.OnBtnPrevClick.bind(this)} className={styles.button} />
            </StackItem>
            <StackItem className={styles.commonstyle}>
            <PrimaryButton text="Next" styles={stackButtonStyles} className={styles.button} onClick={this.SelectSeats.bind(this)} />
             </StackItem><br></br>
             </Stack>
             </Stack>


        }

        {this.state.flag == true && this.state.procflag ==true &&  this.state.Dispalygrid==true &&(this.state.UserLoginName || this.state.Userboolval) &&

        <Stack>

          
        <Stack horizontal tokens={sectionStackTokens1}>

         <StackItem className={styles.commonstyle}>
        <label className={styles.headings}>Buliding Floor Plan :</label>
       
        <Link href={this.state.LinkFloor} target="_blank" data-interception="off">{this.state.MyFloorType}</Link>
        </StackItem>

</Stack> <br/>

<Stack tokens={sectionStackTokens1}>

  <StackItem>

<label className={styles.headings}>Please select the desk(s) required</label>

</StackItem>

</Stack><br/>

<Stack horizontal tokens={sectionStackTokens1} className={styles.MyStyling}>

 {this.state.chckboxesseats.map((item) =>(

   <StackItem>

  
 <Checkbox value={item.seatId} disabled={item.seatTaken} name="chkseats1" label={item.seatId} onChange={this.changechbox.bind(this)} styles={item.seatTaken? checkboxstyles:checkboxstyles} className={item.seatTaken?styles.chkboxDeactive:styles.chkboxActive} >

 </Checkbox>

    
    {/* <input type="checkbox" name="chkseats1" value={item.seatId} disabled={item.seatTaken} className={item.seatTaken? styles.redClass:styles.normalClass} onChange={this.changechbox.bind(this)} />{item.seatId} */}
    
    

    </StackItem>

 ))}

</Stack><br/>


<Stack horizontal tokens={sectionStackTokens1}>

<Checkbox label="Selected Seat" className={styles.resevedSeats1} ></Checkbox>

</Stack><br/>

<Stack horizontal tokens={sectionStackTokens1}>

<Checkbox label="Reserved Seat" className={styles.resevedSeats} ></Checkbox>

</Stack><br/>


<Stack horizontal tokens={sectionStackTokens1}>

<Checkbox label="Empty Seat" className={styles.emptySeats} ></Checkbox>

</Stack><br/>

<Stack horizontal tokens={sectionStackTokens1}>
<label className={styles.headings}>Desk Description of the Seat ID is below:</label>
</Stack><br/>
<Stack horizontal tokens={sectionStackTokens1}>
{this.state.DeskDesc}

</Stack><br/><br/>

<Stack horizontal tokens={sectionStackTokens}>

            <StackItem className={styles.coststyle}>


            <input type="text" name="txtDeskID" value={this.state.deskId} onChange={this.changeDesk.bind(this)}  placeholder="Enter Deskno"/>

            </StackItem>
            <StackItem className={styles.Serachtextbox} id="text">

              <PeoplePicker
                context={this.props.context}
                
                personSelectionLimit={1}
                showtooltip={true}
                required={true}
                disabled={false}
                onChange={this._getPeoplePickerItems}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                defaultSelectedUsers={(this.state.UserLoginName && this.state.UserLoginName.length) ? [this.state.UserLoginName] : []}
                ref={c => (this.ppl = c)} 
                
                //defaultSelectedUsers={['surendra.kalyan.jetty@capco.com']}
                resolveDelay={1000}/>
                
            </StackItem>
            <StackItem>

            <PrimaryButton text="Add" onClick={this.OnAddRowsClick.bind(this)} styles={stackButtonStyles} className={styles.button} />
            </StackItem>

          </Stack>

</Stack>

}
                   

        {this.state.flag == true && this.state.procflag ==true && this.state.GridUserValues && this.state.GridUserValues.length>0 && this.state.Dispalygrid==true &&

          <Stack horizontal tokens={sectionStackTokens}>

           <Grid container className={styles.tableborder}>

           <Grid item md={4}>
              <header className={styles.tablecell}>
              Employee
              </header>

              {this.state.GridUserValues.map((Item,Index)=>(
               <Paper className={styles.tablecelldata1}>{Item.UserName}</Paper>

               ))}

            </Grid>
            <Grid item md={4}>
            <header className={styles.tablecell}>
                Email
              </header>
              {this.state.GridUserValues.map((Item,Index)=>(
              <Paper className={styles.tablecelldata1}>{Item.UserEmail}</Paper>

              ))}
            
            </Grid>
            <Grid item md={2}>
            <header className={styles.tablecell}>
                DeskId
              </header>
              {this.state.GridUserValues.map((Item,Index)=>(
              <Paper className={styles.tablecelldata}>{Item.DeskId}</Paper>

              ))}
            </Grid>
            <Grid item md={2}>
              <header className={styles.tablecell}>
                Remove
              </header>
            {this.state.GridUserValues.map((Item,Index)=>(

<Paper className={styles.tablecelldata}> <IconButton iconProps={{ iconName: "Delete" }} id={Item.DeskId}  className={styles.Iconbutton} onClick={()=>this.onDeleteClick(Item)} /></Paper>


              ))}
            </Grid>


          </Grid>

         </Stack>


        }




{this.state.flag == true && this.state.procflag ==true && this.state.GridUserValues && this.state.GridUserValues.length>0 && this.state.Dispalygrid==true &&

<Stack horizontal tokens={sectionStackTokens}>
  <StackItem>

  <PrimaryButton text="Submit"   styles={stackButtonStyles} className={styles.button} onClick={this.onSubmitClick.bind(this)} />
  </StackItem>

  </Stack>


  }



      </Stack>

    );
  }
}
