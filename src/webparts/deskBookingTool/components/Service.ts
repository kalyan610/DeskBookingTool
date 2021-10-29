import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";


export default class Service {

    public mysitecontext: any;

    public constructor(siteUrl: string, Sitecontext: any) {
        this.mysitecontext = Sitecontext;


        sp.setup({
            sp: {
                baseUrl: siteUrl

            },
        });

    }



   public async GetAllocations():Promise<any>
   {

    
    return await sp.web.lists.getByTitle("ALLCountriesDetails").items.select('Location').expand().get().then(function (data) {

      return data;


    });


   }


   public async GetBulidingData(SelLocVal: string):Promise<any>
   {

    let filtercondition: any = "(Location eq '" + SelLocVal + "')";

    return await  sp.web.lists.getByTitle("ALLCountriesDetails").items.select('BuildingName').filter(filtercondition).get().then(function (data) {

    return data;

    });

   }


   public async GetBookingsData(LocaVal: string, BuildNamVal: string):Promise<any>
   {

    let filtercondition: any = "(Location eq '" + LocaVal + "') and (BuildingName eq '" + BuildNamVal + "')";

    return await sp.web.lists.getByTitle("ALLCountriesDetails").items.select('BookingType').filter(filtercondition).get().then(function (data) {

    return data;

    });

   }



   public async GetFloorsData(LocaVal: string, BuildNamVal: string, BookingTypeval):Promise<any>
   {

    let filtercondition: any = "(Location eq '" + LocaVal + "') and (BuildingName eq '" + BuildNamVal + "') and (BookingType eq '" + BookingTypeval + "')";

    return await sp.web.lists.getByTitle("ALLCountriesDetails").items.select('FloorLevel').filter(filtercondition).get().then(function (data) {

    return data;

    });

   }


    



   public async CheckBlockDate(MyStartDate: string,MyEndDate:string) :Promise<any> {

    const selectedList = 'BlockedDates';
    var BlockStatus='Block';
    let BlockDateexsits=false;
    let filterBlockDates: any = "(Status eq '" + BlockStatus + "') and ((EventDate ge datetime'" + MyStartDate + "' and   EventDate le datetime'" + MyEndDate + "') or (EventDate le datetime'" + MyStartDate + "' and EndDate ge datetime'" + MyStartDate + "' ))";
    try
    {

      
   return await sp.web.lists.getByTitle("BlockedDates").items.select('Title').filter(filterBlockDates).get().then(function (data) {

    for (var k in data) {
           
      if(data[k].Title!='')
      {
        BlockDateexsits=true;
    
      }

      }

      return BlockDateexsits;


   });

  

    }
    catch (error) {
      console.log(error);
  }

}

public async TotalNoofSeats(MyFloorLevel: string,MyBookingType:string) :Promise<any>
{

  let FilterTotalSeats: any = "(BookingType  eq '" + MyBookingType + "') and (Title  eq '" + MyFloorLevel + "')";

   let  TotSeatsistItems = [];

  try
  {
    
 return await sp.web.lists.getByTitle("SeatsList").items.select('Seats').filter(FilterTotalSeats).get().then(function (data) {

  let MySelSeats = data[0].Seats;

  console.log(MySelSeats);

  let arr = MySelSeats.split(',') ;

  for (var k in arr) {


    TotSeatsistItems.push(arr[k]);
    

  }

  return TotSeatsistItems;


 });


  }
  catch (error) {
    console.log(error);
}
  
 
}

public async BookedSeats(MyFloorLevel:string,MyBookingType:string,MyStartDate: string,MyEndDate:string) :Promise<any>
{

  var strbookingsts='cancelled';

  alert(MyFloorLevel);

  let  BookedSeatsListItems = [];

  let FilterBokkedSeats: any = "(BookingStatus  ne '" + strbookingsts + "') and (BookingType eq '" + MyBookingType + "')   and (FloorLevel eq '" + MyFloorLevel + "')  and ((EventDate ge datetime'" + MyStartDate + "' and   EventDate le datetime'" + MyEndDate + "') or (EventDate le datetime'" + MyStartDate + "' and EndDate ge datetime'" + MyStartDate + "' ))";

  try
  {
    
 return await sp.web.lists.getByTitle("ConferenceRoomDetatils").items.select('DeskId').filter(FilterBokkedSeats).get().then(function (data) {


  for(let count=0;count<data.length;count++)
  {

    let MyBookedSeats= data[count].DeskId;

    console.log(MyBookedSeats);
  
    let arr = MyBookedSeats.split(',') ;
  
    for (var k in arr) {
  
  
      BookedSeatsListItems.push(arr[k]);
      
  
    }

  }

 
   
  
  return BookedSeatsListItems;


 });


  }
  catch (error) {
    console.log(error);
}
  
 
}


public async GetUrls(MyUrl: string):Promise<string>
{
 
  let filterBlockDates: any = "(Title eq '" + MyUrl + "')";

  let myrequrl='';
  try
  {
    
 return await sp.web.lists.getByTitle("URLS").items.select('URL').filter(filterBlockDates).get().then(function (data) {

  for (var k in data) {
         
    myrequrl=data[k].URL;

    }

    return myrequrl;


 });


  }
  catch (error) {
    console.log(error);
}


}

//Latest

    public async MyGetAllocations():Promise<any>
   {

    
    return await sp.web.lists.getByTitle("Locations").items.select('Title').get().then(function (data) {

      return data;


    });


   }


   public async MyGetBulidingData(SelLocVal: string):Promise<any>
   {

    let filtercondition: any = "(Title eq '" + SelLocVal + "')";

    return await  sp.web.lists.getByTitle("CountriesDetails").items.select('BulidingName').filter(filtercondition).get().then(function (data) {

    return data;

    });

   }


   public async MyGetBookingType(LocaVal: string, BuildNamVal: string):Promise<any>
   {

    let filtercondition: any = "(Title eq '" + LocaVal + "') and (BuildingName eq '" + BuildNamVal + "')";

    return await sp.web.lists.getByTitle("BookingInformation").items.select('BookingType').filter(filtercondition).get().then(function (data) {

    return data;

    });

   }

   public async MyGetFloorsData(LocaVal: string, BuildNamVal: string, BookingTypeval):Promise<any>
   {

    let filtercondition: any = "(Title eq '" + LocaVal + "') and (BuildingName eq '" + BuildNamVal + "') and (BookingType eq '" + BookingTypeval + "')";

    return await sp.web.lists.getByTitle("BookingandFloorDetails").items.select('FloorLevel').filter(filtercondition).get().then(function (data) {

    return data;

    });

   }




   public async getCurrentUserSiteGroups(): Promise<any[]> {

    try {

        return (await sp.web.currentUser.groups.select("Id,Title,Description,OwnerTitle,OnlyAllowMembersViewMembership,AllowMembersEditMembership,Owner/Id,Owner/LoginName").expand('Owner').get());

    }
    catch {
        throw 'get current user site groups failed.';
    }

}


public async getCurrentUser(): Promise<any> {
  try {
      return await sp.web.currentUser.get().then(result => {
          return result;
      });
  } catch (error) {
      console.log(error);
  }
}


//END



private async Save(MyLocation:string,MyBuildingName:string,MyBookingType:string,MyFloorLevel:string,MyStartDate: string,MyEndDate:string,MYDeskId:string,MyEmail:string,MyTitle:string):Promise<any>     {       
  
  await sp.web.lists.getByTitle('ConferenceRoomDetatils').items.add({       
    
    Location:MyLocation,
    BuildingName:MyBuildingName,
    BookingType:MyBookingType,
    FloorLevel:MyFloorLevel,
    EventDate:MyStartDate,            
    EndDate:MyEndDate,
    NumStatus:'2',
    BookingStatus:'Booked',
    DeskId:MYDeskId,
    Email:MyEmail,
    Title:MyTitle
  
  });

}




public async GetPDFLinks1(MyBuilding:string,MyBookingType:string,MyFloorLevel:string):Promise<any>
{
  
  let filtercondition: any = " (Building eq '" + MyBuilding + "') and (BookingType eq '" + MyBookingType + "') and (Floor eq '" + MyFloorLevel + "')" ;

   return await sp.web.lists.getByTitle("FloorPlans").items.select('URL').filter(filtercondition).get().then(function (data) {

    return data[0].URL;
  
  });

}


public async GetDeskDesc(MyDeskID:string):Promise<any>
{

   let filtercondition: any = " (SeatNumber eq '" + MyDeskID + "')" ;

  return await sp.web.lists.getByTitle("SeatsDescription").items.select('SeatNumber','Description').filter(filtercondition).get().then(function (data) {

  return data[0].Description;
  
  });

}


}