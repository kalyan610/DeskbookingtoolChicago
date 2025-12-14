import * as React from 'react';
import styles from './DeskbookingTool2025Version2.module.scss';
import type { IDeskbookingTool2025Version2Props } from './IDeskbookingTool2025Version2Props';
import { DateTimePicker,TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { Grid, Paper} from '@material-ui/core';
import Service from './Service'


import {
 
  Stack,IStackStyles,Dropdown,IDropdownStyles,PrimaryButton,IStackTokens,StackItem,IDropdownOption,
  Checkbox,ICheckboxStyles,IconButton
  
} from '@fluentui/react';


//let RootUrl = '';

let Userreqemail = '';
let UserreqName = '';
let TotalSeatsList = [];
let BookedSeatsList:any = [];

const stackStyles: Partial<IStackStyles> = { root: { padding: 10 } };
const sectionStackTokens1: IStackTokens = { childrenGap: 5 };

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },

};




const dropdownStyles1: Partial<IDropdownStyles> = {
  dropdown: { width: 150 },
 

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


  let AllCheckedItems: any = [];

  
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


  const stackButtonStyles: Partial<IStackStyles> = { root: { Width: 20 } };

  //const stackButtonStyles1: Partial<IStackStyles> = { root: { Width: 60 } };
const stackTokens = { childrenGap: 80 };

const sectionStackTokens: IStackTokens = { childrenGap: 10 };





   export interface IDeskbookingtool2025Version2 {
    Mylocationval: any;
    Mylocationval1:any;
    MyBookingTypeVal: any;
    MyFloorType: any;
    MyFloorType1:any;
    flag: boolean;
    procflag:boolean;
    startDate?: Date
    EndDate: any;
    BlockCount: number;
    ItemInfo: any;
    ConcString: string;
    LocationListItems: any;
    LocationListItems1:any;
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
    //ChangeQuestions:any;
    QuestKey:any;
    isDisable:boolean;
    FirstDivVisble:boolean;
    HideSecondPage:boolean;
    HideThirdPage:boolean;
    
    
  
  }

  interface ILocationItem {
  key: string;
  text: string;
}

  interface IBuldingItems {
  key: string;
  text: string;
}

  interface IFloorLevels {
  key: string;
  text: string;
}

 interface IBookingTypes {
  key: string;
  text: string;
}

     

export default class DeskbookingTool2025Version2 extends React.Component<IDeskbookingTool2025Version2Props,IDeskbookingtool2025Version2> 
{

  public _service: any;
  public GlobalService: any;
  protected ppl:any;

  public constructor(props: IDeskbookingTool2025Version2Props) {
  super(props);
   this.state = {
  
        Mylocationval: null,
        Mylocationval1:null,
       MyBookingTypeVal: null,
        MyFloorType: null,
        MyFloorType1:null,
        flag: false,
        procflag:false,
        NoofSeats: null,
        startDate: undefined,
        EndDate: null,
        BlockCount: 0,
        ItemInfo: null,
        ConcString: "",
        LocationListItems: [],
       LocationListItems1:[],
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
       FirstDivVisble:true,
       HideSecondPage:false,
       HideThirdPage:false
       
  
      };


   
      
      //RootUrl = this.props.url;
  
    this._service = new Service(this.props.url, this.props.context);
  
      this.getAllLocations();
      
     this.getAllLocations1();
      this.getAllFloorLevels();
       this.getAllBookingTypes();

    }


  public async getAllLocations() {

    alert('one');

     let myLocationLocal: ILocationItem[] = [];

  const data: any[] = await this._service.MyGetAllocations();

  console.log(data);

  let AllLocations: ILocationItem[] = data.map(item => ({
    key: item.Title,
    text: item.Title
  }));

  console.log(AllLocations);

  AllLocations.forEach((item: ILocationItem) => {
    const itemExists = myLocationLocal.some(ditem => ditem.key === item.key);

    if (!itemExists) {
      myLocationLocal.push({ key: item.key, text: item.text });
    }
  });

  this.setState({ LocationListItems: myLocationLocal,Mylocationval: AllLocations[0].text });

  }


   public async getAllLocations1() {

    alert('one');

     let myLocationLocal1: ILocationItem[] = [];

  const data: any[] = await this._service.MyGetAllocations1();

  console.log(data);

  let AllLocations: ILocationItem[] = data.map(item => ({
    key: item.Title,
    text: item.Title
  }));

  console.log(AllLocations);

  AllLocations.forEach((item: ILocationItem) => {
    const itemExists = myLocationLocal1.some(ditem => ditem.key === item.key);

    if (!itemExists) {
      myLocationLocal1.push({ key: item.key, text: item.text });
    }
  });

  this.setState({ LocationListItems1: myLocationLocal1,Mylocationval1: AllLocations[0].text });

  }


    public async getAllFloorLevels() {

     let myFloorLevelLocal: IFloorLevels[] = [];

  const data: any[] = await this._service.MyGetAllFloorLevels();

  console.log(data);

  let AllFloorLevels: IFloorLevels[] = data.map(item => ({
    key: item.Title,
    text: item.Title
  }));

  console.log(AllFloorLevels);

  AllFloorLevels.forEach((item: IBuldingItems) => {
    const itemExists = myFloorLevelLocal.some(ditem => ditem.key === item.key);

    if (!itemExists) {
      myFloorLevelLocal.push({ key: item.key, text: item.text });
    }
  });

  this.setState({ FloorListItems: myFloorLevelLocal, MyFloorType:AllFloorLevels[0].text});

  }

  public async getAllBookingTypes() {

     let myBookingTypeLocal: IBookingTypes[] = [];

  const data: any[] = await this._service.MyGetAllBookinTypes();

  console.log(data);

  let AllBookingTypes: IBookingTypes[] = data.map(item => ({
    key: item.Title,
    text: item.Title
  }));

  console.log(AllBookingTypes);

  AllBookingTypes.forEach((item: IBookingTypes) => {
    const itemExists = myBookingTypeLocal.some(ditem => ditem.key === item.key);

    if (!itemExists) {
      myBookingTypeLocal.push({ key: item.key, text: item.text });
    }
  });

  this.setState({ BookingListItems: myBookingTypeLocal });

  }


   private handleChangeLocation(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void 
   {

    
    this.setState({ Mylocationval: item.text });

  }

     private handleChangeLocation1(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void 
   {

    
    this.setState({ Mylocationval1: item.text });

  }



  private handleChangeBookingType(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {


    
    this.setState({ MyBookingTypeVal: item.text });
    

  }


  private handleChangeFloorLevel(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {

    
    this.setState({ MyFloorType: item.text });


  }

  private _onFormatDate = (date: Date): string => {
      return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();

     
  };

     private ChangeSeats(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption):void{
        this.setState({NoofSeats:item});
      }

     public formatdate(strDate: Date | undefined): string {
  if (!strDate) return "";       // handles undefined case
  return strDate.toISOString();  // since it's already a Date
}


  public checkItemInArray(itemName: string, ItemArray: any): boolean {
    let isItemExist = false;

    for (var bkseatcount = 0; bkseatcount < ItemArray.length; bkseatcount++) {
      if (ItemArray[bkseatcount] === itemName)
        isItemExist = true;
    }
    return isItemExist;
  }


       public async SelectSeats() {


    let MyStartDate = this.formatdate(this.state.startDate);

    let MyEndDate = this.formatdate(this.state.EndDate);

    let MyFloorLevel = this.state.MyFloorType;

    let MyBookingType = this.state.MyBookingTypeVal;

    if (this.state.NoofSeats == null || this.state.NoofSeats == '' || this.state.NoofSeats=='Select') {

      alert('please select No of Seats');

    }


    else if (this.state.startDate == null || this.state.EndDate == '') {

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

      //this.GlobalService = new Service(this.props.url, this.props.context);


      //let MyUrl = await this.GlobalService.GetUrls(this.state.Mylocationval);

      //alert(MyUrl);

      this._service = new Service(this.props.url, this.props.context);

      let BlockStatus = await this._service.CheckBlockDate(MyStartDate, MyEndDate);

     


      if (BlockStatus == true) {


        alert("Date has been blocked");

      }

      if (BlockStatus == false) {


          this.setState({HideThirdPage:true})
           this.setState({HideSecondPage:true})


        //var MyBuild=this.state.MyBuildingVal;
        //let  MyHyperlink= await this._service.GetPDFLinks1(MyBuild,MyBookingType,MyFloorLevel);
        //this.setState({LinkFloor:MyHyperlink});

        

        TotalSeatsList = await this._service.TotalNoofSeats(MyFloorLevel, MyBookingType);

        BookedSeatsList = await this._service.BookedSeats(MyFloorLevel, MyBookingType, MyStartDate, MyEndDate);

        if(BookedSeatsList!=null)
        {
         
          //var strhtml = '';
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

          //var strhtml = '';
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
          this.setState({HideThirdPage:true})
           this.setState({HideSecondPage:true})

        }

      }

    }

  }


  // private changeDesk(data: any): void {

  //   this.setState({ deskId: data.target.value });

  // }


    private async changechbox(data: any) {


     //New Lines

     console.log(data.target.id);

     

     //let SelDeskId=data.target.attributes["aria-label"].id;

     let SelDeskId=data.target.id;

     let Checkstatusval=data.target.checked;

    //  this.GlobalService = new Service(this.props.url, this.props.context);

    // let MyUrl2 = await this.GlobalService.GetUrls(this.state.Mylocationval);

    this._service = new Service(this.props.url, this.props.context);

    //Gan
    
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

    private async _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);

    if(items.length>0)
    {

    Userreqemail = items[0].secondaryText;
    UserreqName = items[0].text

    }

    //this.ppl.onChange([]);

  }

   private changeDesk(data: any): void {

    this.setState({ deskId: data.target.value });

  }

private OnAddRowsClick(): boolean {
  // Validate Desk ID
  if (!this.state.deskId) {
    alert('Please enter the Desk ID');
    return false;
  }

  // Validate User email
  if (!Userreqemail) {
    alert('Please select a username');
    return false;
  }

  // Check selected items
  const isItemChecked = this.checkItemChecked(this.state.deskId, AllCheckedItems);
  const isUserNameExists = this.checkItemUserName(Userreqemail, this.state.GridUserValues);
  const isDeskIDExists = this.checkDeskIDchked(this.state.deskId, this.state.GridUserValues);

  if (!isItemChecked) {
    alert('Please enter the desk number that you selected above');
    return false;
  }

  if (isUserNameExists) {
    alert('Please select a different user name');
    return false;
  }

  if (isDeskIDExists) {
    alert('Please enter a different desk ID');
    return false;
  }

  // If valid data → Add to the grid
  if (isItemChecked && !isUserNameExists && !isDeskIDExists) {
    const gridLocal = [...(this.state.GridUserValues || [])];

    const gridItem = {
      UserEmail: Userreqemail,
      UserName: UserreqName,
      DeskId: this.state.deskId
    };

    gridLocal.push(gridItem);

    this.setState({
      GridUserValues: gridLocal,
      deskId: '',
      Userboolval: true,
      MyUserName: '',
      checkstatus: false
    });

    Userreqemail = '';
    UserreqName = '';
    this.ppl.onChange([]);  // Reset People Picker
  }

  return true; // Success condition
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


  private OnBtnClick(): void {

     if (this.state.Mylocationval == null || this.state.Mylocationval == 'Select Location' || this.state.Mylocationval == 'Select') {

      alert('please select location');
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
       this.setState({ HideSecondPage: true });
    }

   }

   private OnBtnPrevClick(): void {

    this.setState({ flag: false });

    this.setState({Dispalygrid:false});

    this.setState({GridUserValues:[]});

    this.setState({ HideSecondPage: false });

    this.setState({ HideThirdPage: false });

    //AllCheckedItems=[];

     //Added Code

     


  }


  private async onSubmitClick()
  {

    
    //this.GlobalService = new Service(this.props.url, this.props.context);

    //let MyUrl1 = await this.GlobalService.GetUrls(this.state.Mylocationval);

    this._service = new Service(this.props.url, this.props.context);

    
    for(let count=0;count<this.state.GridUserValues.length;count++)
    {

      let MyTilte=this.state.Mylocationval + "-" +this.state.GridUserValues[count].UserName + "-" +this.state.MyFloorType 

      this._service.onDrop(this.state.Mylocationval,'Wills Tower',this.state.MyBookingTypeVal,this.state.MyFloorType,this.state.startDate,this.state.EndDate,this.state.GridUserValues[count].DeskId,this.state.GridUserValues[count].UserEmail,MyTilte).then(function (data:any)
      {

     
     window.location.replace("https://capcoinc.sharepoint.com/sites/ChicagoDeskReservation");

      });

     

    }

    alert('Submitted Successfully');

    
  }

     
public render(): React.ReactElement<IDeskbookingTool2025Version2Props> {

    return (
    
<Stack tokens={stackTokens} styles={stackStyles}>
      
{this.state.HideSecondPage==false && this.state.HideThirdPage==false &&
<Stack>
      <label className={styles.headings}>Location</label>
            <br></br>
            <Dropdown options={this.state.LocationListItems} className={styles.headings} styles={dropdownStyles} selectedKey={this.state.Mylocationval ? this.state.Mylocationval : undefined} onChange={this.handleChangeLocation.bind(this)} /><br></br>
                        <label className={styles.headings}>Building</label><br></br>
            <Dropdown  options={this.state.LocationListItems1} styles={dropdownStyles} selectedKey={this.state.Mylocationval1 ? this.state.Mylocationval1 : undefined} onChange={this.handleChangeLocation1.bind(this)} /><br></br>
            
        <label className={styles.headings}>Floor Level</label><br></br>
        <Dropdown  options={this.state.FloorListItems} styles={dropdownStyles} selectedKey={this.state.MyFloorType ? this.state.MyFloorType : undefined} onChange={this.handleChangeFloorLevel.bind(this)} /><br></br>
           
           
            <label className={styles.headings}>Booking Type</label><br></br>
            <Dropdown placeHolder="Select Booking Type" options={this.state.BookingListItems} styles={dropdownStyles} selectedKey={this.state.MyBookingTypeVal ? this.state.MyBookingTypeVal : undefined} onChange={this.handleChangeBookingType.bind(this)} /><br></br>

            <PrimaryButton text="Next" onClick={this.OnBtnClick.bind(this)} styles={stackButtonStyles} className={styles.button} />
      
</Stack>

}

{this.state.HideSecondPage==true &&

<Stack className={styles.dateTimeClass}>
<div>
  <label className={styles.headings}>Start Date</label><br></br>
  </div><br></br>
<DateTimePicker
        value={this.state.startDate}
        onChange={(date) => this.setState({ startDate: date })}
  timeDisplayControlType={TimeDisplayControlType.Dropdown}
  formatDate={this._onFormatDate}
  /><br></br>
<div>
  <label className={styles.headings}>End Date</label><br></br>
  </div><br></br>
<DateTimePicker
        value={this.state.EndDate}
        onChange={(date) => this.setState({ EndDate: date })}
  timeDisplayControlType={TimeDisplayControlType.Dropdown}
  formatDate={this._onFormatDate}
  /><br></br>

<Stack horizontal tokens={sectionStackTokens}>
 <StackItem className={styles.commonstyle}>

              <label className={styles.headings}>Number of Seats:</label><br></br>
               
              </StackItem>
              <StackItem className={styles.commonstyle}>
              <Dropdown
                placeholder="Select No of seats"
                options={options}
                styles={dropdownStyles1}
                selectedKey={this.state.NoofSeats ? this.state.NoofSeats.key : undefined}
                onChange={this.ChangeSeats.bind(this)} 
            />   

            </StackItem>
</Stack>
<br></br>
<br></br>

<Stack>
<Stack horizontal tokens={sectionStackTokens}>
            <StackItem className={styles.commonstyle}>
            <PrimaryButton text="Previous" styles={stackButtonStyles}  onClick={this.OnBtnPrevClick.bind(this)} className={styles.button} />
            </StackItem>
            <StackItem className={styles.commonstyle}>
            <PrimaryButton text="Start Selecting" styles={stackButtonStyles} className={styles.button1} onClick={this.SelectSeats.bind(this)} />
             </StackItem><br></br>
</Stack>
</Stack>

</Stack>
 
 }

{this.state.HideSecondPage==true && this.state.HideThirdPage==true &&

<Stack>


<Stack tokens={sectionStackTokens1}>

  <StackItem>

<label className={styles.headings}>Please select the desk(s) required</label>

</StackItem>

</Stack><br/>

<Stack horizontal tokens={sectionStackTokens1} className={styles.MyStyling}>

 {this.state.chckboxesseats.map((item:any) =>(

  <StackItem>
  
 <Checkbox id={item.seatId} disabled={item.seatTaken} name="chkseats1" label={item.seatId} onChange={this.changechbox.bind(this)} styles={item.seatTaken? checkboxstyles:checkboxstyles} className={item.seatTaken?styles.chkboxDeactive:styles.chkboxActive} >

 </Checkbox>

 </StackItem>

 ))}

</Stack><br/>


<Stack horizontal tokens={sectionStackTokens1}>

<Checkbox label="Selected Seat" className={styles.resevedSeats1}></Checkbox>

</Stack><br/>

<Stack horizontal tokens={sectionStackTokens1}>

<Checkbox label="Reserved Seat" className={styles.resevedSeats} ></Checkbox>

</Stack><br/>

<Stack horizontal tokens={sectionStackTokens1}>

<Checkbox label="Empty Seat" className={styles.emptySeats} ></Checkbox>

</Stack><br/> 

<Stack tokens={{ childrenGap: 4 }} className={styles.labelBlock}>
  <label className={styles.headings}>Desk Description</label><br></br>
  <label className={styles.resevedSeats}>{this.state.DeskDesc}</label>
</Stack>
<br>
</br>

{/* <div className={styles.cardContainer}> */}

<h3 className={styles.sectionTitle}>Assign Desk</h3>

  {/* <Stack horizontal tokens={{ childrenGap: 12 }}  className={styles.inputRow}> */}
     <Stack horizontal tokens={{ childrenGap: 12 }}>
    
    {/* Desk Input */}
    {/* className={styles.formItemSpacing} */}
    <StackItem className={styles.formItemSpacing}>
      <input
        type="text"
        name="txtDeskID"
        value={this.state.deskId}
        onChange={this.changeDesk.bind(this)}
        placeholder="Enter Desk Number"
        className={styles.inputBox}
        />
    </StackItem>

    {/* PeoplePicker */}
    <StackItem className={styles.formItemSpacing}>
     
      <PeoplePicker 
        context={this.props.context}
        personSelectionLimit={1}
        showtooltip={true}
        required={true}
        disabled={false}
        placeholder='Please enter Email'
        onChange={this._getPeoplePickerItems}
        principalTypes={[PrincipalType.User, PrincipalType.SecurityGroup]}
        webAbsoluteUrl="https://capcoinc.sharepoint.com/"
        defaultSelectedUsers={this.state.UserLoginName ? [this.state.UserLoginName] : []}
        resolveDelay={1000}
       ref={(c) => (this.ppl = c)}
      />
      
    </StackItem>

    {/* Add Button */}
    <StackItem>
      <PrimaryButton 
        text="Add" 
        className={styles.button}
        styles={stackButtonStyles}
        onClick={this.OnAddRowsClick.bind(this)} 
      />
     </StackItem>
    
  </Stack>
{/* </div> */}

<br></br>
<Stack>

  <Stack horizontal tokens={sectionStackTokens}>

           <Grid container className={styles.tableborder}>

           <Grid item md={4}>
              <header className={styles.tablecell}>
              Employee
              </header>

              {this.state.GridUserValues.map((Item:any,Index:any)=>(
               <Paper className={styles.tablecelldata1}>{Item.UserName}</Paper>

               ))}

            </Grid>
            <Grid item md={4}>
            <header className={styles.tablecell}>
                Email
              </header>
              {this.state.GridUserValues.map((Item:any,Index:any)=>(
              <Paper className={styles.tablecelldata1}>{Item.UserEmail}</Paper>

              ))}
            
            </Grid>
            <Grid item md={2}>
            <header className={styles.tablecell}>
                DeskId
              </header>
              {this.state.GridUserValues.map((Item:any,Index:any)=>(
              <Paper className={styles.tablecelldata}>{Item.DeskId}</Paper>

              ))}
            </Grid>
            <Grid item md={2}>
              <header className={styles.tablecell}>
                Remove
              </header>
            {this.state.GridUserValues.map((Item:any,Index:any)=>(

<Paper className={styles.tablecelldata}> <IconButton iconProps={{ iconName: "Delete" }} id={Item.DeskId}  className={styles.Iconbutton} onClick={()=>this.onDeleteClick(Item)} /></Paper>


              ))}
            </Grid>


          </Grid>

  </Stack>

  

</Stack>

</Stack>

}

{this.state.HideSecondPage==true &&  this.state.HideThirdPage==true &&

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
