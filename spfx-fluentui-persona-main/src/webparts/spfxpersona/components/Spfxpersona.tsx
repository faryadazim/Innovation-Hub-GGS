import * as React from 'react';
import styles from './Spfxpersona.module.scss';
import { ISpfxpersonaProps } from './ISpfxpersonaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp/presets/all";
import { RenderProfilePicture } from '../Common/Components/RenderProfilePicture/RenderProfilePicture';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import * as _ from 'lodash';

export interface ISpfxpersonaWebPartState {
  ProductionBoardDevelopers: any[];
}

export default class Spfxpersona extends React.Component<ISpfxpersonaProps, ISpfxpersonaWebPartState> {

  constructor(props: ISpfxpersonaProps) {
    super(props);
    this.state = {
      ProductionBoardDevelopers: []
    };
  }

  private getUserProfileUrl = async (loginName: string) => {
    const userPictureUrl = await sp.profiles.getUserProfilePropertyFor(loginName, 'PictureURL');
    return userPictureUrl;
  }

  private getNextWeekDates(): { startDate: string; endDate: string } {
    const currentDate = new Date();
    const currentDayOfWeek = currentDate.getDay(); // 0: Sunday, 1: Monday, ..., 6: Saturday

    // Calculate the number of days until the next Monday (start of the next week)
    const daysUntilNextMonday = currentDayOfWeek === 0 ? 1 : 8 - currentDayOfWeek;

    // Calculate the start and end dates of the next week
    const nextWeekStartDate = new Date();
    nextWeekStartDate.setDate(currentDate.getDate() + daysUntilNextMonday);

    const nextWeekEndDate = new Date();
    nextWeekEndDate.setDate(nextWeekStartDate.getDate() + 6); // Adding 6 days to get the end of the week (Sunday)

    let StartDateOfWeekWithStartTime: string =
      (nextWeekStartDate.getFullYear()) +
      '-' +
      ((nextWeekStartDate.getMonth() + 1) < 10 ? '0' + (nextWeekStartDate.getMonth() + 1) : nextWeekStartDate.getMonth() + 1) +
      '-' +
      (nextWeekStartDate.getDate() < 10 ? '0' + nextWeekStartDate.getDate() : nextWeekStartDate.getDate());
    StartDateOfWeekWithStartTime += "T00:00:00Z";
    let EndDateOfWeekWithEndTime: string =
      (nextWeekEndDate.getFullYear()) +
      '-' +
      ((nextWeekEndDate.getMonth() + 1) < 10 ? '0' + (nextWeekEndDate.getMonth() + 1) : nextWeekEndDate.getMonth() + 1) +
      '-' +
      (nextWeekEndDate.getDate() < 10 ? '0' + nextWeekEndDate.getDate() : nextWeekEndDate.getDate());
    EndDateOfWeekWithEndTime += "T23:59:59Z";

    // Return the start and end dates of the next week
    return { startDate: StartDateOfWeekWithStartTime, endDate: EndDateOfWeekWithEndTime };
  }

  private getCurrentWeekDates(): { startDate: string; endDate: string } {
    const today = new Date();
    const currentDayOfWeek = today.getDay(); // 0 (Sunday) to 6 (Saturday)

    // Calculate the date of the start of the week (Monday)
    const startOfWeek = new Date();
    startOfWeek.setDate(today.getDate() - currentDayOfWeek + 1);

    // Calculate the date of the end of the week (Sunday)
    const endOfWeek = new Date();
    endOfWeek.setDate(today.getDate() - currentDayOfWeek + 7);

    let StartDateOfWeekWithStartTime: string =
      (startOfWeek.getFullYear()) +
      '-' +
      ((startOfWeek.getMonth() + 1) < 10 ? '0' + (startOfWeek.getMonth() + 1) : startOfWeek.getMonth() + 1) +
      '-' +
      (startOfWeek.getDate() < 10 ? '0' + startOfWeek.getDate() : startOfWeek.getDate());
    StartDateOfWeekWithStartTime += "T00:00:00Z";
    let EndDateOfWeekWithEndTime: string =
      (endOfWeek.getFullYear()) +
      '-' +
      ((endOfWeek.getMonth() + 1) < 10 ? '0' + (endOfWeek.getMonth() + 1) : endOfWeek.getMonth() + 1) +
      '-' +
      (endOfWeek.getDate() < 10 ? '0' + endOfWeek.getDate() : endOfWeek.getDate());
    EndDateOfWeekWithEndTime += "T23:59:59Z";

    // Return the start and end dates of the current week
    return {
      startDate: StartDateOfWeekWithStartTime,
      endDate: EndDateOfWeekWithEndTime,
    };
  }

  private getNextWeekNumberAndYear(): { weekNumber: number, year: number } {
    const currentDate = new Date();
    const currentDayOfWeek = currentDate.getDay(); // 0: Sunday, 1: Monday, ..., 6: Saturday

    // Calculate the number of days until the next Monday (start of the next week)
    const daysUntilNextMonday = currentDayOfWeek === 0 ? 1 : 8 - currentDayOfWeek;

    // Calculate the start and end dates of the next week
    const nextWeekStartDate = new Date();
    nextWeekStartDate.setDate(currentDate.getDate() + daysUntilNextMonday);

    const year = nextWeekStartDate.getFullYear();

    const startDate = new Date(nextWeekStartDate.getFullYear(), 0, 1);
    const days = Math.floor(((nextWeekStartDate as any) - (startDate as any)) / (24 * 60 * 60 * 1000));

    const weekNumber = Math.ceil(days / 7);

    return { weekNumber: weekNumber, year: year };
  }

  public async getProductionBoardDevelopers() {
    // ProductionBoard
    // ActivityProductionBoard
    // GET /_api/web/lists/getbytitle('YourFirstList')/items?$filter=currentWeekDates.start ge datetime'yyyy-MM-ddT00:00:00' and currentWeekDates.end lt datetime'yyyy-MM-ddT23:59:59'&$select=Id,Title,OtherFields

    // const nextWeekDates = this.getNextWeekDates();
    // console.log("Start of the next week:", nextWeekDates.startDate);
    // console.log("End of the next week:", nextWeekDates.endDate);

    // const currentWeekDates = this.getCurrentWeekDates();
    // console.log("Start of the current week:", currentWeekDates.startDate);
    // console.log("End of the current week:", currentWeekDates.endDate);

    // Getting the week number of the next week
    const nextWeekInfo = this.getNextWeekNumberAndYear();
    console.log("Next week's info:", nextWeekInfo);
    
    let siteUsers = {};
    (await sp.web.siteUsers.get()).forEach((userInfo: ISiteUserInfo, index: number, array: ISiteUserInfo[]) => {
      siteUsers[userInfo.Id] = { Title: userInfo.Title, LoginName: userInfo.LoginName };
    });
    // console.log("siteUsers: " + siteUsers);

    let camelQueryXML: string =
      '<View>' +
      "<ViewFields>" +
      "<FieldRef Name='ID'/>" +
      "<FieldRef Name='Title'/>" +
      "<FieldRef Name='Developer' LookupId='TRUE'/>" +
      "<FieldRef Name='Week'/>" +
      "<FieldRef Name='Year'/>" +
      "</ViewFields>" +
      "<Query>" +
      "<OrderBy>" +
      "<FieldRef Name='ID' Ascending='TRUE'/>" +
      "</OrderBy>" +
      "<Where>" +
      "<And>" +
      "<Eq>" +
      "<FieldRef Name='Week' />" +
      "<Value Type='Number'>" + nextWeekInfo.weekNumber + "</Value>" +
      "</Eq>" +
      "<Eq>" +
      "<FieldRef Name='Year' />" +
      "<Value Type='Number'>" + nextWeekInfo.year + "</Value>" +
      "</Eq>" +
      "</And>" +
      "</Where>" +
      "</Query>" +
      '</View>';

    //const productionBoardData = await sp.web.lists.getByTitle("ProductionBoard").items.filter("Created ge '" + currentWeekDates.start + "' and Created le '" + currentWeekDates.end + "'").select("Title", "Developer/Title", "Developer/ID").expand("Developer").top(100).get();
    //const activityProductionBoardData = await sp.web.lists.getByTitle("ActivityProductionBoard").items.select("Title", "Developer/Title", "Developer/ID").expand("Developer").getAll();

    sp.web.lists.getByTitle("ProductionBoard").getItemsByCAMLQuery({ 'ViewXml': camelQueryXML }).then((productionBoardResponse: any) => {
      // console.log("productionBoardData: " + productionBoardResponse);
      // Use Set to store distinct items based on Person or Group field Id
      const pbDistinctItemsSet = new Set<number>();
      const distinctProductionBoardResponse = productionBoardResponse.filter((item: any) => {
        if (!pbDistinctItemsSet.has(item.DeveloperId)) {
          pbDistinctItemsSet.add(item.DeveloperId);
          return true;
        }
        return false;
      });
      sp.web.lists.getByTitle("ActivityProductionBoard").getItemsByCAMLQuery({ 'ViewXml': camelQueryXML }).then((activityProductionBoardResponse: any) => {
        // console.log("productionBoardData: " + activityProductionBoardResponse);

        // Use Set to store distinct items based on Person or Group field Id
        const apbDistinctItemsSet = new Set<number>();
        const distinctActivityProductionBoardResponse = activityProductionBoardResponse.filter((item: any) => {
          if (!apbDistinctItemsSet.has(item.DeveloperId)) {
            apbDistinctItemsSet.add(item.DeveloperId);
            return true;
          }
          return false;
        });

        if (distinctProductionBoardResponse.length > 0 && distinctActivityProductionBoardResponse.length > 0) {
          // const productionBoardfilteredData: any = _.difference(distinctProductionBoardResponse, distinctActivityProductionBoardResponse);
          // const productionBoardfilteredData = distinctProductionBoardResponse.filter((pbItem: any) => distinctActivityProductionBoardResponse.some((apbItem: any) => apbItem.DeveloperId !== pbItem.DeveloperId));

          const productionBoardfilteredData = distinctProductionBoardResponse.filter((element1: any) => {
            return !distinctActivityProductionBoardResponse.find((element2: any) => {
              return element2.DeveloperId === element1.DeveloperId;
            });
          });

          productionBoardfilteredData.forEach(element => {
            element.DeveloperName = siteUsers[element.DeveloperId].Title;
            element.LoginName = siteUsers[element.DeveloperId].LoginName;
          });

          this.setState({ ProductionBoardDevelopers: productionBoardfilteredData });
        }
        else if (distinctProductionBoardResponse.length > 0) {
          distinctProductionBoardResponse.forEach(element => {
            element.DeveloperName = siteUsers[element.DeveloperId].Title;
            element.LoginName = siteUsers[element.DeveloperId].LoginName;
          });

          this.setState({ ProductionBoardDevelopers: distinctProductionBoardResponse });
        }
      }).catch(error => {
        console.log("Error while getting ProductionBoard data from list", error);
      });
    }).catch(error => {
      console.log("Error while getting ProductionBoard data from list", error);
    });
  }

  public componentDidMount() {
    this.getProductionBoardDevelopers();
  }

  public render(): React.ReactElement<ISpfxpersonaProps> {
    return (
      <div className={styles.spfxpersona}>
        <span><b>Production Board</b></span>
        <br></br>
        <span>Loaded for next week</span>
        {this.state.ProductionBoardDevelopers.length == 0 ?
          <div>No records found!</div> :
          this.state.ProductionBoardDevelopers.map(developer =>
            <RenderProfilePicture
              developerName={developer.DeveloperName}
              title={developer.Title}
              getUserProfileUrl={() => this.getUserProfileUrl(developer.LoginName)}  ></RenderProfilePicture>
          )}
      </div>
    );
  }
}
