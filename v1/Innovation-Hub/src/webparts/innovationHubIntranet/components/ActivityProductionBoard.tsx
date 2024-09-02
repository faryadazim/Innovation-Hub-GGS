import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from "@pnp/sp/webs";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  PrimaryButton,
  TextField,
  ITextFieldStyles,
  Spinner,
  ILabelStyles,
  Toggle,
  Modal,
  NormalPeoplePicker,
  TooltipHost,
  TooltipOverflowMode,
  IColumn,
  DatePicker,
  IDatePickerStyles,
} from "@fluentui/react";

import Service from "../components/Services";

import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
import { FontIcon, IIconProps } from "@fluentui/react/lib/Icon";
import "../ExternalRef/styleSheets/Styles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import Pagination from "office-ui-fabric-react-pagination";
import { IDetailsListStyles } from "office-ui-fabric-react";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

//Sorting
let sortApbDataArr = [];
let sortApbFilterArr = [];
let sortApbUpdate = false;

let gblATPDetails = [];
let FilterProject;
let FilterProduct;
let ProjectOrProductDetails = [];
let DateListFormat = "DD/MM/YYYY";
let DatePickerFormat = "YYYY-MM-DDT14:00:00Z";

function ActivityProductionBoard(props: any) {
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;
  let pbSwitchID = props.pbSwitch ? props.pbSwitch.split("-")[0] : null;
  let pbSwitchType = props.pbSwitch ? props.pbSwitch.split("-")[1] : null;

  let Apb_ActivityPlanId = props.ActivityPlanID;
  let navType = props.pageType;

  let Apb_Year = moment().year();
  let Apb_NextWeekYear = moment().add(1, "week").year();
  let Apb_LastWeekYear = moment().subtract(1, "week").year();

  let Apb_WeekNumber = moment().isoWeek();
  let Apb_NextWeekNumber = moment().add(1, "week").isoWeek();
  let Apb_LastWeekNumber = moment().subtract(1, "week").isoWeek();

  let thisWeekMonday = moment().day(1).format("YYYY-MM-DD");
  let thisWeekTuesday = moment().day(2).format("YYYY-MM-DD");
  let thisWeekWednesday = moment().day(3).format("YYYY-MM-DD");
  let thisWeekThursday = moment().day(4).format("YYYY-MM-DD");
  let thisWeekFriday = moment().day(5).format("YYYY-MM-DD");
  let thisWeekSaturday = moment().day(6).format("YYYY-MM-DD");
  let thisWeekSunday = moment().day(7).format("YYYY-MM-DD");

  let loggeduseremail = props.spcontext.pageContext.user.email;
  // let loggeduseremail = "contract.developer@goodtogreatschools.org.au";
  let currentpage = 1;
  let totalPageItems = 10;
  const allPeoples = props.peopleList;
  let loggeduserid = allPeoples.filter(
    (dev) => dev.secondaryText == loggeduseremail
  )[0].ID;
  let loggerusername = allPeoples.filter(
    (dev) => dev.secondaryText == loggeduseremail
  )[0].text;

  // Initialization function
  const drAllitems = {
    Request: null,
    Requestto: null,
    Emailcc: null,
    Project: null,
    Documenttype: null,
    Link: null,
    Comments: null,
    Confidential: false,
    IsExternalAllow: false,
    Product: null,
    AnnualPlanID: null,
    DeliveryPlanID: null,
    ProductionBoardID: null,
  };

  const ApbFilterKeys = {
    Lessons: "All",
    Steps: "All",
    Product: "All",
    Project: "All",
    Showonly: "Mine",
    WeekNumber: Apb_WeekNumber,
    Year: Apb_Year,
    Week: "This Week",
    Developer: loggeduserid,
  };
  let AdrPBErrorStatus = {
    Request: "",
    Requestto: "",
    Documenttype: "",
    Link: "",
  };
  let ApbErrorStatus = {
    Type: "",
    StartDate: "",
    EndDate: "",
    Project: "",
    Product: "",
    Lessons: "",
    Steps: "",
    PlannedHours: "",
  };
  const ApbDrpDwnOptns = {
    Lessons: [{ key: "All", text: "All" }],
    Steps: [{ key: "All", text: "All" }],
    Product: [{ key: "All", text: "All" }],
    Project: [{ key: "All", text: "All" }],
    Showonly: [
      { key: "Mine", text: "Mine" },
      { key: "All", text: "All" },
    ],
    Week: [
      { key: "This Week", text: "This Week" },
      { key: "Last Week", text: "Last Week" },
      { key: "Next Week", text: "Next Week" },
    ],
    WeekNumber: [{ key: Apb_WeekNumber, text: Apb_WeekNumber.toString() }],
    Year: [{ key: Apb_Year, text: Apb_Year.toString() }],
    DeveloperMine: [{ key: loggeduserid, text: loggerusername }],
    Developer: [{ key: loggeduserid, text: loggerusername }],
  };
  const ApbModalBoxDrpDwnOptns = {
    Request: [],
    Documenttype: [],
    Type: [],
    Project: [],
    Product: [],
  };

  //Detail list Columns
  const _apbColumns = [
    {
      key: "Column1",
      name: "Type",
      fieldName: "Title",
      minWidth: 60,
      maxWidth: 60,
      onColumnClick: (ev, column) => {
        !sortApbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Title}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Title}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column2",
      name: "Start date",
      fieldName: "StartDate",
      minWidth: 80,
      maxWidth: 80,
      onColumnClick: (ev, column) => {
        !sortApbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => item.StartDate,
    },
    {
      key: "Column3",
      name: "End date",
      fieldName: "EndDate",
      minWidth: 75,
      maxWidth: 75,
      onColumnClick: (ev, column) => {
        !sortApbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => item.EndDate,
    },
    {
      key: "Column4",
      name: "Source",
      fieldName: "Source",
      minWidth: 60,
      maxWidth: 60,
      onColumnClick: (ev, column) => {
        !sortApbUpdate ? _onColumnClick(ev, column) : null;
      },
    },
    {
      key: "APName",
      name: "AP name",
      fieldName: "AP name",
      minWidth: 120,
      maxWidth: 200,
      onRender: (item) => {
        let curAPName = gblATPDetails.filter((arr) => {
          return arr.ID == item.ActivityPlanID;
        });

        return curAPName.length > 0 ? curAPName[0].ActivityPlanName : "";
      },
    },
    {
      key: "Column5",
      name: "Name of the deliverable",
      fieldName: "Project",
      minWidth: 120,
      maxWidth: 200,
      onColumnClick: (ev, column) => {
        !sortApbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Project}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Project}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column6",
      name: "Product",
      fieldName: "Product",
      minWidth: 120,
      maxWidth: 200,
      onColumnClick: (ev, column) => {
        !sortApbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Product}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Product}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column7",
      name: "Section",
      fieldName: "Lessons",
      minWidth: 80,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        !sortApbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Lessons}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Lessons}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column8",
      name: "Steps",
      fieldName: "Steps",
      minWidth: 120,
      maxWidth: 250,
      onColumnClick: (ev, column) => {
        !sortApbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Steps}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Steps}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column9",
      name: "PH/UH",
      fieldName: "PlannedHours",
      minWidth: 60,
      maxWidth: 60,
      onColumnClick: (ev, column) => {
        !sortApbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item, index: number) =>
      // item.PHWeek ? Math.round(item.PlannedHours) + "Wks" : item.PlannedHours,
      {
        if (item.PHWeek) {
          let valPH = item.PlannedHours.toString();
          valPH = valPH.split(".");
          let resultPH;
          if (valPH.length > 1) {
            if (valPH[0] == "0") {
              resultPH =
                Math.round((item.PlannedHours - valPH[0]) * 7) + " D ";
            } else {
              resultPH =
                Math.round(valPH[0]) +
                " W " +
                Math.round((item.PlannedHours - valPH[0]) * 7) +
                " D ";
            }
          } else {
            resultPH = Math.round(item.PlannedHours) + "W";
          }
          return resultPH;
        } else {
          return (
            <>
              {item.UnPlannedHours ? (
                <span style={{ color: "#FAA332", fontWeight: 600 }}>
                  {item.PlannedHours}
                </span>
              ) : (
                <span style={{ color: "#038387", fontWeight: 600 }}>
                  {item.PlannedHours}
                </span>
              )}
            </>
          );
        }
      },
    },
    {
      key: "Column10",
      name: "Mon",
      fieldName: "Monday",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) => {
        let thisWeekMonday = moment()
          .isoWeek(item.Week)
          .day(1)
          .format("YYYY-MM-DD");
        return (
          <TextField
            styles={{
              root: {
                selectors: {
                  ".ms-TextField-fieldGroup": {
                    borderRadius: 4,
                    border: "1px solid",
                    height: 28,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              ApbUpdate &&
                item.DeveloperEmail == loggeduseremail &&
                thisWeekMonday >=
                moment(item.StartDate, DateListFormat).format("YYYY-MM-DD") &&
                thisWeekMonday <=
                moment(item.EndDate, DateListFormat).format("YYYY-MM-DD")
                ? false
                : true
            }
            value={item.Monday}
            onChange={(e: any) => {
              // parseFloat(e.target.value)
              //   ? ApbOnchangeItems(item.RefId, "Monday", e.target.value)
              //   : ApbOnchangeItems(item.RefId, "Monday", null);
              ApbOnchangeItems(item.RefId, "Monday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column11",
      name: "Tue",
      fieldName: "Tuesday",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) => {
        let thisWeekTuesday = moment()
          .isoWeek(item.Week)
          .day(2)
          .format("YYYY-MM-DD");
        return (
          <TextField
            styles={{
              root: {
                selectors: {
                  ".ms-TextField-fieldGroup": {
                    borderRadius: 4,
                    border: "1px solid",
                    height: 28,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              ApbUpdate &&
                item.DeveloperEmail == loggeduseremail &&
                thisWeekTuesday >=
                moment(item.StartDate, DateListFormat).format("YYYY-MM-DD") &&
                thisWeekTuesday <=
                moment(item.EndDate, DateListFormat).format("YYYY-MM-DD")
                ? false
                : true
            }
            value={item.Tuesday}
            onChange={(e: any) => {
              // parseFloat(e.target.value)
              //   ? ApbOnchangeItems(item.RefId, "Tuesday", e.target.value)
              //   : ApbOnchangeItems(item.RefId, "Tuesday", null);
              ApbOnchangeItems(item.RefId, "Tuesday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column12",
      name: "Wed",
      fieldName: "Wednesday",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) => {
        let thisWeekWednesday = moment()
          .isoWeek(item.Week)
          .day(3)
          .format("YYYY-MM-DD");
        return (
          <TextField
            styles={{
              root: {
                selectors: {
                  ".ms-TextField-fieldGroup": {
                    borderRadius: 4,
                    border: "1px solid",
                    height: 28,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              ApbUpdate &&
                item.DeveloperEmail == loggeduseremail &&
                thisWeekWednesday >=
                moment(item.StartDate, DateListFormat).format("YYYY-MM-DD") &&
                thisWeekWednesday <=
                moment(item.EndDate, DateListFormat).format("YYYY-MM-DD")
                ? false
                : true
            }
            value={item.Wednesday}
            onChange={(e: any) => {
              // parseFloat(e.target.value)
              //   ? ApbOnchangeItems(item.RefId, "Wednesday", e.target.value)
              //   : ApbOnchangeItems(item.RefId, "Wednesday", null);
              ApbOnchangeItems(item.RefId, "Wednesday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column13",
      name: "Thu",
      fieldName: "Thursday",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) => {
        let thisWeekThursday = moment()
          .isoWeek(item.Week)
          .day(4)
          .format("YYYY-MM-DD");
        return (
          <TextField
            styles={{
              root: {
                selectors: {
                  ".ms-TextField-fieldGroup": {
                    borderRadius: 4,
                    border: "1px solid",
                    height: 28,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              ApbUpdate &&
                item.DeveloperEmail == loggeduseremail &&
                thisWeekThursday >=
                moment(item.StartDate, DateListFormat).format("YYYY-MM-DD") &&
                thisWeekThursday <=
                moment(item.EndDate, DateListFormat).format("YYYY-MM-DD")
                ? false
                : true
            }
            value={item.Thursday}
            onChange={(e: any) => {
              // parseFloat(e.target.value)
              //   ? ApbOnchangeItems(item.RefId, "Thursday", e.target.value)
              //   : ApbOnchangeItems(item.RefId, "Thursday", null);
              ApbOnchangeItems(item.RefId, "Thursday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column14",
      name: "Fri",
      fieldName: "Friday",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) => {
        let thisWeekFriday = moment()
          .isoWeek(item.Week)
          .day(5)
          .format("YYYY-MM-DD");

        return (
          <TextField
            styles={{
              root: {
                selectors: {
                  ".ms-TextField-fieldGroup": {
                    borderRadius: 4,
                    border: "1px solid",
                    height: 28,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              ApbUpdate &&
                item.DeveloperEmail == loggeduseremail &&
                thisWeekFriday >=
                moment(item.StartDate, DateListFormat).format("YYYY-MM-DD") &&
                thisWeekFriday <=
                moment(item.EndDate, DateListFormat).format("YYYY-MM-DD")
                ? false
                : true
            }
            value={item.Friday}
            onChange={(e: any) => {
              // parseFloat(e.target.value)
              //   ? ApbOnchangeItems(item.RefId, "Friday", e.target.value)
              //   : ApbOnchangeItems(item.RefId, "Friday", null);
              ApbOnchangeItems(item.RefId, "Friday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column15",
      name: "Sat",
      fieldName: "Saturday",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) => {
        let thisWeekSaturday = moment()
          .isoWeek(item.Week)
          .day(6)
          .format("YYYY-MM-DD");

        return (
          <TextField
            styles={{
              root: {
                selectors: {
                  ".ms-TextField-fieldGroup": {
                    borderRadius: 4,
                    border: "1px solid",
                    height: 28,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              ApbUpdate &&
                item.DeveloperEmail == loggeduseremail &&
                thisWeekSaturday >=
                moment(item.StartDate, DateListFormat).format("YYYY-MM-DD") &&
                thisWeekSaturday <=
                moment(item.EndDate, DateListFormat).format("YYYY-MM-DD")
                ? false
                : true
            }
            value={item.Saturday}
            onChange={(e: any) => {
              
              ApbOnchangeItems(item.RefId, "Saturday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column15",
      name: "Sun",
      fieldName: "Sunday",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) => {
        let thisWeekSunday = moment()
          .isoWeek(item.Week)
          .day(7)
          .format("YYYY-MM-DD");

        return (
          <TextField
            styles={{
              root: {
                selectors: {
                  ".ms-TextField-fieldGroup": {
                    borderRadius: 4,
                    border: "1px solid",
                    height: 28,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              ApbUpdate &&
                item.DeveloperEmail == loggeduseremail &&
                thisWeekSunday >=
                moment(item.StartDate, DateListFormat).format("YYYY-MM-DD") &&
                thisWeekSunday <=
                moment(item.EndDate, DateListFormat).format("YYYY-MM-DD")
                ? false
                : true
            }
            value={item.Sunday}
            onChange={(e: any) => {
              ApbOnchangeItems(item.RefId, "Sunday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column16",
      name: "AH",
      fieldName: "ActualHours",
      minWidth: 40,
      maxWidth: 40,
      onColumnClick: (ev, column) => {
        !sortApbUpdate ? _onColumnClick(ev, column) : null;
      },
    },
    {
      key: "Column17",
      name: "Action",
      fieldName: "",
      minWidth: 65,
      maxWidth: 65,
      onRender: (item, Index) =>
        // item.Week == Apb_WeekNumber &&
        item.DeveloperEmail == loggeduseremail && item.ID != 0 ? (
          <div
            style={{
              marginTop: "-6px",
            }}
          >
            <Icon
              iconName="OpenEnrollment"
              title={item.Status}
              style={{
                color:
                  item.Status == null
                    ? "#0882A5"
                    : item.Status == "Pending"
                      ? "#000000"
                      : item.Status == "Signed Off" ||
                        item.Status == "Published" ||
                        item.Status == "Publish ready" ||
                        item.Status == "Completed"
                        ? "#40b200"
                        : item.Status == "Returned" || item.Status == "Cancelled"
                          ? "#ff3838"
                          : "#ffb302",
                marginTop: 6,
                fontSize: 17,
                height: 14,
                width: 17,
                cursor: "pointer",
              }}
              onClick={(_) => {
                drAllitems.Project = item.Project;
                drAllitems.Product = item.Product;
                drAllitems.AnnualPlanID = item.ActivityPlanID;
                drAllitems.DeliveryPlanID = item.ActivityDeliveryPlanID;
                drAllitems.ProductionBoardID = item.ID;
                setApbButtonLoader(false);
                setAdrPBShowMessage(AdrPBErrorStatus);
                setApbDocumentReview(drAllitems);
                setApbModalBoxVisibility(true);
              }}
            />
            {item.Source == "Ad hoc" ? (
              <>
                <Icon
                  iconName="Edit"
                  title="Edit deliverable"
                  className={ApbiconStyleClass.edit}
                  onClick={() => {
                    setApbButtonLoader(false);
                    let adhocItem = {
                      RefId: item.RefId,
                      ID: item.ID,
                      StartDate: new Date(
                        moment(item.StartDate, DateListFormat).format(
                          DatePickerFormat
                        )
                      ),
                      EndDate: new Date(
                        moment(item.EndDate, DateListFormat).format(
                          DatePickerFormat
                        )
                      ),
                      Source: item.Source,
                      Project: item.Project,
                      Product: item.Product,
                      Title: item.Title,
                      PlannedHours: item.PlannedHours,
                      Monday: item.Monday,
                      Tuesday: item.Tuesday,
                      Wednesday: item.Wednesday,
                      Thursday: item.Thursday,
                      Friday: item.Friday,
                      Saturday: item.Saturday,
                      Sunday: item.Sunday,
                      ActualHours: item.ActualHours,
                      DeveloperId: item.DeveloperId,
                      DeveloperEmail: item.DeveloperEmail,
                      Week: item.Week,
                      Year: item.Year,
                      Status: item.Status,
                      Lessons: item.Lessons,
                      Steps: item.Steps,
                      ActivityPlanID: item.ActivityPlanID,
                      ActivityDeliveryPlanID: item.ActivityDeliveryPlanID,
                      ADPActualHours: item.ADPActualHours,
                      UnPlannedHours: item.UnPlannedHours,
                      PHWeek: item.PHWeek,
                      Onchange: item.Onchange,
                    };
                    setApbShowMessage(ApbErrorStatus);
                    setApbAdhocPopup({
                      visible: true,
                      isNew: false,
                      value: adhocItem,
                    });
                  }}
                />
                <Icon
                  iconName="Delete"
                  title="Delete deliverable"
                  className={ApbiconStyleClass.delete}
                  onClick={() => {
                    setApbButtonLoader(false),
                      setApbDeletePopup({ condition: true, targetId: item.ID });
                  }}
                />
              </>
            ) : (
              ""
            )}
          </div>
        ) : item.ID != 0 ? (
          <Icon
            iconName="OpenEnrollment"
            title={item.Status}
            style={{
              color:
                item.Status == null
                  ? "#0882A5"
                  : item.Status == "Pending"
                    ? "#000000"
                    : item.Status == "Signed Off" ||
                      item.Status == "Published" ||
                      item.Status == "Publish ready" ||
                      item.Status == "Completed"
                      ? "#40b200"
                      : item.Status == "Returned" || item.Status == "Cancelled"
                        ? "#ff3838"
                        : "#ffb302",
              marginTop: 6,
              fontSize: 17,
              height: 14,
              width: 17,
              cursor: "pointer",
            }}
            onClick={(_) => { }}
          />
        ) : (
          <Icon
            iconName="OpenEnrollment"
            style={{
              color: "#ababab",
              marginTop: 6,
              fontSize: 17,
              height: 14,
              width: 17,
              cursor: "default",
            }}
            onClick={(_) => { }}
          />
        ),
    },
  ];

  // Design
  const saveIcon: IIconProps = { iconName: "Save" };
  const editIcon: IIconProps = { iconName: "Edit" };
  const cancelIcon: IIconProps = { iconName: "Cancel" };
  const ApbModalBoxDatePickerStyles: Partial<IDatePickerStyles> = {
    root: {
      width: "300px",
      margin: "10px 20px",
      borderRadius: "4px",
    },
    icon: {
      fontSize: "17px",
      color: "#000",
      fontWeight: "bold",
    },
    textField: {
      selectors: {
        ".ms-TextField-fieldGroup": {
          height: "36px",
        },
      },
    },
  };
  const dateFormater = (date: Date): string => {
    return date ? moment(date).format("DD/MM/YYYY") : "";
  };
  const gridStyles: Partial<IDetailsListStyles> = {
    root: {
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          ".ms-DetailsRow-fields": {
            alignItems: "center",
          },
        },
      },
    },
    headerWrapper: {
      flex: "0 0 auto",
    },
    contentWrapper: {
      flex: "1 1 auto",
      overflowY: "auto",
      overflowX: "hidden",
    },
  };
  const ApbLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const ApbProjectInfo = mergeStyles({
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    margin: "0 10px",
  });
  const ApblabelStyles = mergeStyleSets({
    titleLabel: [
      {
        color: "#676767",
        fontSize: "14px",
        marginRight: "10px",
        fontWeight: "400",
      },
    ],
    selectedLabel: [
      {
        color: "#0882A5",
        fontSize: "14px",
        marginRight: "10px",
        fontWeight: "600",
      },
    ],
    labelValue: [
      {
        color: "#0882A5",
        fontSize: "14px",
        marginRight: "10px",
      },
    ],
    inputLabels: [
      {
        color: "#323130",
        fontSize: "13px",
      },
    ],
    ErrorLabel: [
      {
        marginTop: "25px",
        marginLeft: "10px",
        fontWeight: "500",
        color: "#D0342C",
        fontSize: "13px",
      },
    ],
    NORLabel: [
      {
        color: "#323130",
        fontSize: "13px",
        marginLeft: "10px",
        fontWeight: "500",
      },
    ],
  });
  const ApbBigiconStyleClass = mergeStyleSets({
    ChevronLeftMed: [
      {
        cursor: "pointer",
        color: "#2392b2",
        fontSize: 24,
        marginTop: "3px",
        marginRight: 12,
      },
    ],
  });
  const ApbbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const ApbbuttonStyleClass = mergeStyleSets({
    buttonPrimary: [
      {
        color: "White",
        backgroundColor: "#FAA332",
        borderRadius: "3px",
        border: "none",
        marginRight: "10px",
        selectors: {
          ":hover": {
            backgroundColor: "#FAA332",
            opacity: 0.9,
            borderRadius: "3px",
            border: "none",
            marginRight: "10px",
          },
        },
      },
      ApbbuttonStyle,
    ],
    buttonSecondary: [
      {
        color: "White",
        backgroundColor: "#038387",
        borderRadius: "3px",
        border: "none",
        margin: "0 5px",
        selectors: {
          ":hover": {
            backgroundColor: "#038387",
            opacity: 0.9,
          },
        },
      },
      ApbbuttonStyle,
    ],
  });
  const ApbiconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const ApbiconStyleClass = mergeStyleSets({
    link: [{ color: "blue", margin: "0 0" }, ApbiconStyle],
    delete: [{ color: "#CB1E06", margin: "0 0px" }, ApbiconStyle],
    edit: [{ color: "#2392B2", margin: "0 7px" }, ApbiconStyle],
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 22,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginTop: 27,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    Apblink: [
      {
        color: "white",
        fontSize: "18px",
        height: 22,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginLeft: 10,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      },
    ],
    export: [
      {
        color: "black",
        fontSize: "18px",
        height: 20,
        width: 20,
        cursor: "pointer",
        marginRight: 5,
      },
    ],
  });
  const ApbDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };

  const ApbActiveShortDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 75,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: 4,
      fontWeight: 600,
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
    callout: {
      maxHeight: 300,
    },
  };
  const ApbShortLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 75,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };

  const ApbActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };

  const showonlyDropdown: Partial<IDropdownStyles> = {
    root: {
      width: 70,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const showonlyDropdownActive: Partial<IDropdownStyles> = {
    root: {
      width: 70,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const ApbModalBoxDropdownStyles: Partial<IDropdownStyles> = {
    root: { width: "300px", margin: "10px 20px" },
    title: {
      fontSize: 12,
      borderRadius: "4px",
      border: "1px solid",
      height: "36px",
      padding: "3px 10px",
      color: "#000",
    },
    dropdownItemsWrapper: { fontSize: 12 },
    dropdownItemSelected: { fontSize: 12 },
    caretDown: {
      fontSize: 14,
      padding: "3px",
      color: "#000",
      fontWeight: "bold",
    },
  };
  const ApbModalBoxDrpDwnCalloutStyles: Partial<IDropdownStyles> = {
    root: { width: "300px", margin: "10px 20px" },
    title: {
      fontSize: 12,
      borderRadius: "4px",
      border: "1px solid",
      padding: "3px 10px",
      height: "36px",
      color: "#000",
    },
    dropdownItemsWrapper: { fontSize: 12 },
    dropdownItemSelected: { fontSize: 12 },
    caretDown: {
      fontSize: 14,
      paddingTop: "3px",
      color: "#000",
      fontWeight: "bold",
    },
    callout: { height: 200 },
  };
  const ApbTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "300px",
      margin: "10px 20px",
    },
    field: {
      fontSize: 12,
      color: "#000",
      borderRadius: "4px",
      background: "#fff !important",
    },
    fieldGroup: {
      border: "1px solid !important",
      height: "36px",
    },
  };
  const ApbMultiTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "640px",
      margin: "10px 20px",
      borderRadius: "4px",
    },
    field: { fontSize: 12, color: "#000" },
  };
  const ApbModalBoxPP = mergeStyles({
    width: "300px",
    margin: "10px 20px",
  });

  // useState
  const [ApbReRender, setApbReRender] = useState(false);
  const [ApbChecked, setApbChecked] = useState(false);
  const [ApbUpdate, setApbUpdate] = useState(false);
  const [ApbDisplayData, setApbDisplayData] = useState([]);
  const [ApbFilterData, setApbFilterData] = useState([]);
  const [ApbData, setApbData] = useState([]);
  const [ApbMasterData, setApbMasterData] = useState([]);
  const [ApbDropDownOptions, setApbDropDownOptions] = useState(ApbDrpDwnOptns);
  const [ApbFilterOptions, setApbFilterOptions] = useState(ApbFilterKeys);
  const [ApbcurrentPage, setApbCurrentPage] = useState(currentpage);
  const [ApbLoader, setApbLoader] = useState(true);
  const [ApbModalBoxVisibility, setApbModalBoxVisibility] = useState(false);
  const [ApbButtonLoader, setApbButtonLoader] = useState(false);
  const [ApbModalBoxDropDownOptions, setApbModalBoxDropDownOptions] = useState(
    ApbModalBoxDrpDwnOptns
  );
  const [ApbDocumentReview, setApbDocumentReview] = useState(drAllitems);
  const [AdrPBShowMessage, setAdrPBShowMessage] = useState(AdrPBErrorStatus);
  const [ApbShowMessage, setApbShowMessage] = useState(ApbErrorStatus);
  const [ApbWeek, setApbWeek] = useState(Apb_WeekNumber);
  const [ApbYear, setApbYear] = useState(Apb_Year);
  // const [ApbLastweek, setApbLastweek] = useState([]);
  // const [ApbNextweek, setApbNextweek] = useState([]);
  const [ApbAutoSave, setApbAutoSave] = useState(false);
  const [apbColumns, setapbColumns] = useState(_apbColumns);

  const [documentLinkStatus, setDocumentLinkStatus] = useState("no-checked")
  const [ApbAdhocPopup, setApbAdhocPopup] = useState({
    visible: false,
    isNew: true,
    value: {},
  });
  const [ApbDeletePopup, setApbDeletePopup] = useState({
    condition: false,
    targetId: 0,
  });
  // useEffect
  useEffect(() => {
    getModalBoxOptions();
    getATPDetails();
    // Apb_ActivityPlanId
    //   ? getCurrentApbData(Apb_WeekNumber, Apb_Year, ApbFilterKeys)
    //   : getApbData(loggeduserid, Apb_WeekNumber, Apb_Year, ApbFilterKeys);
  }, [ApbReRender]);

  useEffect(() => {
    if (ApbAutoSave && ApbUpdate) {
      setTimeout(() => {
        document.getElementById("apdBtnSave").click();
      }, 300000);
    }
  }, [ApbAutoSave]);

  window.onbeforeunload = function (e) {
    debugger;
    if (ApbAutoSave) {
      let dialogText =
        "You have unsaved changes, are you sure you want to leave?";
      e.returnValue = dialogText;
      return dialogText;
    }
  };

  const alertDialogforBack = () => {
    if (confirm("You have unsaved changes, are you sure you want to leave?")) {
      navType == "ATP"
        ? props.handleclick("ActivityPlan")
        : props.handleclick("ActivityDeliveryPlan", Apb_ActivityPlanId);
    }
  };

  // Functions
  const getModalBoxOptions = () => {
    const _sortFilterKeys = (a, b) => {
      if (a.text.toLowerCase() < b.text.toLowerCase()) {
        return -1;
      }
      if (a.text.toLowerCase() > b.text.toLowerCase()) {
        return 1;
      }
      return 0;
    };

    //Request Choices
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .fields.getByInternalNameOrTitle("Request")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              ApbModalBoxDrpDwnOptns.Request.findIndex((rApb) => {
                return rApb.key == choice;
              }) == -1
            ) {
              ApbModalBoxDrpDwnOptns.Request.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch((error) => {
        ApbErrorFunction(error, "getModalBoxOptions1");
      });

    //Documenttype Choices
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .fields.getByInternalNameOrTitle("Documenttype")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              ApbModalBoxDrpDwnOptns.Documenttype.findIndex((rdt) => {
                return rdt.key == choice;
              }) == -1
            ) {
              ApbModalBoxDrpDwnOptns.Documenttype.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch((error) => {
        ApbErrorFunction(error, "getModalBoxOptions2");
      });

    //type Choices
    sharepointWeb.lists
      .getByTitle("Product List")
      .fields.getByInternalNameOrTitle("Types")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              ApbModalBoxDrpDwnOptns.Type.findIndex((rdt) => {
                return rdt.key == choice;
              }) == -1
            ) {
              ApbModalBoxDrpDwnOptns.Type.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch((error) => {
        ApbErrorFunction(error, "getModalBoxOptions3");
      });

    //Product Choices
    let NotSureId = null;
    sharepointWeb.lists
      .getByTitle("Master Product List")
      .items.filter("IsDeleted ne 1")
      .top(5000)
      .get()
      .then((allProducts) => {
        allProducts.forEach((product) => {
          if (product.Title != null) {
            if (
              ApbModalBoxDropDownOptions.Product.findIndex((productOptn) => {
                return productOptn.key == product.Title;
              }) == -1
            ) {
              if (product.Title != "Not Sure") {
                ApbModalBoxDropDownOptions.Product.push({
                  key: product.Title + " " + product.ProductVersion,
                  text: product.Title + " " + product.ProductVersion,
                });
              }
              ProjectOrProductDetails.push({
                Type: "Product",
                Id: product.ID,
                Key: product.Title + " " + product.ProductVersion,
                Title: product.Title,
                Version: product.ProductVersion,
              });
            }
          }
        });
      })
      .then(() => {
        ApbModalBoxDropDownOptions.Product.sort(_sortFilterKeys);
        ApbModalBoxDropDownOptions.Product.unshift({
          key: "Not Sure V1",
          text: "Not Sure V1",
        });
      })
      .catch((error) => {
        ApbErrorFunction(error, "getModalBoxOptions4");
      });

    //Project Choices
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.top(5000)
      .get()
      .then((allPrj) => {
        allPrj.forEach((prj) => {
          if (prj.Title != null) {
            if (
              ApbModalBoxDropDownOptions.Project.findIndex((productOptn) => {
                return productOptn.key == prj.Title;
              }) == -1
            ) {
              ApbModalBoxDropDownOptions.Project.push({
                key: prj.Title + " " + prj.ProjectVersion,
                text: prj.Title + " " + prj.ProjectVersion,
              });
              ProjectOrProductDetails.push({
                Type: "Project",
                Id: prj.ID,
                Key: prj.Title + " " + prj.ProjectVersion,
                Title: prj.Title,
                Version: prj.ProjectVersion,
              });
            }
          }
        });
      })
      .then(() => {
        ApbModalBoxDropDownOptions.Project.sort(_sortFilterKeys);
      })
      .catch((error) => {
        ApbErrorFunction(error, "getModalBoxOptions5");
      });

    setApbModalBoxDropDownOptions(ApbModalBoxDrpDwnOptns);
  };
  const generateExcel = () => {
    let arrExport = ApbFilterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Type", key: "Type", width: 25 },
      { header: "Start date", key: "Startdate", width: 25 },
      { header: "End date", key: "Enddate", width: 25 },
      { header: "Source", key: "Source", width: 25 },
      { header: "Name of the deliverable", key: "POT", width: 25 },
      { header: "Product", key: "Product", width: 60 },
      { header: "Section", key: "Lessons", width: 20 },
      { header: "Steps", key: "Steps", width: 20 },
      {
        header: "Planned hours/Unplanned hours",
        key: "PlannedHours",
        width: 40,
      },
      { header: "Monday", key: "Monday", width: 30 },
      { header: "Tuesday", key: "Tuesday", width: 30 },
      { header: "Wednesday", key: "Wednesday", width: 30 },
      { header: "Thursday", key: "Thursday", width: 30 },
      { header: "Friday", key: "Friday", width: 30 },
      { header: "Saturday", key: "Saturday", width: 30 },
      { header: "Sunday", key: "Sunday", width: 30 },
      { header: "Actual hours", key: "ActualTotal", width: 30 },
    ];
    arrExport.forEach((item) => {
      worksheet.addRow({
        Type: item.Title ? item.Title : "",
        Startdate: item.StartDate ? item.StartDate : "",
        Enddate: item.EndDate ? item.EndDate : "",
        Source: item.Source ? item.Source : "",
        POT: item.Project ? item.Project : "",
        Product: item.Product ? item.Product : "",
        Lessons: item.Lessons ? item.Lessons : "",
        Steps: item.Steps ? item.Steps : "",
        PlannedHours: item.PlannedHours ? item.PlannedHours : "",
        Monday: item.Monday ? item.Monday : "",
        Tuesday: item.Tuesday ? item.Tuesday : "",
        Wednesday: item.Wednesday ? item.Wednesday : "",
        Thursday: item.Thursday ? item.Thursday : "",
        Friday: item.Friday ? item.Friday : "",
        Saturday: item.Saturday ? item.Saturday : "",
        Sunday: item.Sunday ? item.Sunday : "",

        ActualTotal: item.ActualHours ? item.ActualHours : "",
      });
    });
    [
      "A1",
      "B1",
      "C1",
      "D1",
      "E1",
      "F1",
      "G1",
      "H1",
      "I1",
      "J1",
      "K1",
      "L1",
      "M1",
      "N1",
      "O1",
    ].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    [
      "A1",
      "B1",
      "C1",
      "D1",
      "E1",
      "F1",
      "G1",
      "H1",
      "I1",
      "J1",
      "K1",
      "L1",
      "M1",
      "N1",
      "O1",
    ].map((key) => {
      worksheet.getCell(key).color = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF" },
      };
    });
    workbook.xlsx
      .writeBuffer()
      .then((buffer) =>
        FileSaver.saveAs(
          new Blob([buffer]),
          `ActivityProductionBoard-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };

  const getATPDetails = () => {
    gblATPDetails = [];
    sharepointWeb.lists
      .getByTitle("Activity Plan")
      .items.select("*", "FieldValuesAsText/CompletedDate")
      .expand("FieldValuesAsText")
      .orderBy("Modified", false)
      .top(5000)
      .get()
      .then((items) => {
        items.forEach((item) => {
          gblATPDetails.push({
            ID: item.Id ? item.Id : "",
            Project: item.Project
              ? item.Project +
              " " +
              (item.ProjectVersion ? item.ProjectVersion : "V1")
              : "",
            ActivityPlanName: item.ActivityPlanName
              ? item.ActivityPlanName
              : "",
            Product: item.Product
              ? item.Product +
              " " +
              (item.ProductVersion ? item.ProductVersion : "V1")
              : "",
          });
        });
        Apb_ActivityPlanId
          ? getCurrentApbData(Apb_WeekNumber, Apb_Year, ApbFilterKeys)
          : getApbData(loggeduserid, Apb_WeekNumber, Apb_Year, ApbFilterKeys);
      })
      .catch((err) => {
        console.log(err);
      });
  };

  const getCurrentApbData = (Week, Year, filterkeys) => {
    setApbLoader(true);
    sharepointWeb.lists
      .getByTitle("ActivityProductionBoard")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
      )
      .expand("Developer,FieldValuesAsText")
      .filter(
        "ActivityPlanID eq '" +
        Apb_ActivityPlanId +
        "' and Week eq '" +
        Week +
        "' and Year eq '" +
        Year +
        "' "
      )
      .top(5000)
      .get()
      .then((items) => {
        let _ApbAllitems = [];
        let Count = 0;
        console.log(items, "activty production board");
        items.forEach(async (item, Index) => {
          let curATPDetails = gblATPDetails.filter((arr) => {
            return arr.ID == item.ActivityPlanID;
          });

          // For onchange calculation
          let oldProduct =
            item.Product +
            " " +
            (item.ProductVersion ? item.ProductVersion : "V1");
          let oldProject =
            item.Project +
            " " +
            (item.ProjectVersion ? item.ProjectVersion : "V1");
          let NewProject =
            curATPDetails.length > 0 ? curATPDetails[0].Project : "";
          let NewProduct =
            curATPDetails.length > 0 ? curATPDetails[0].Product : "";
          // const reponseDate = await getCamelquery(item.ActivityDeliveryPlanID);
          // let reponseDateActuall = moment(
          //   item["FieldValuesAsText"].EndDate,
          //   DateListFormat
          // ).format(DateListFormat);


          // if (reponseDate) {
          //   reponseDateActuall = reponseDate
          // }
          _ApbAllitems.push({
            RefId: Index + 1,
            ID: item.ID,
            StartDate: moment(
              item["FieldValuesAsText"].StartDate,
              DateListFormat
            ).format(DateListFormat),
            EndDate: moment(
              item["FieldValuesAsText"].EndDate,
              DateListFormat
            ).format(DateListFormat),
            Source: item.Source,
            Project:
              curATPDetails.length > 0
                ? curATPDetails[0].Project
                : item.Project
                  ? item.Project +
                  " " +
                  (item.ProjectVersion ? item.ProjectVersion : "V1")
                  : "",
            Product:
              curATPDetails.length > 0
                ? curATPDetails[0].Product
                : item.Product
                  ? item.Product +
                  " " +
                  (item.ProductVersion ? item.ProductVersion : "V1")
                  : "",
            Title: item.Title,
            PlannedHours: item.PlannedHours ? item.PlannedHours : 0,
            Monday: item.Monday ? item.Monday : "0",
            Tuesday: item.Tuesday ? item.Tuesday : "0",
            Wednesday: item.Wednesday ? item.Wednesday : "0",
            Thursday: item.Thursday ? item.Thursday : "0",
            Friday: item.Friday ? item.Friday : "0",
            Saturday: item.Saturday ? item.Saturday : "0",
            Sunday: item.Sunday ? item.Sunday : "0",
            ActualHours: item.ActualHours,
            DeveloperId: item.DeveloperId,
            DeveloperEmail: item.Developer ? item.Developer.EMail : "",
            Week: item.Week,
            Year: item.Year,
            Status: item.Status,
            Lessons: item.Lessons,
            Steps: item.Steps,
            ActivityPlanID: item.ActivityPlanID,
            ActivityDeliveryPlanID: item.ActivityDeliveryPlanID,
            ADPActualHours: item.ADPActualHours ? item.ADPActualHours : 0,
            UnPlannedHours: item.UnPlannedHours ? item.UnPlannedHours : false,
            PHWeek: item.PHWeek ? item.PHWeek : null,
            Onchange:
              oldProduct != NewProduct || oldProject != NewProject
                ? true
                : false,
          });
        });

        // if (_ApbAllitems.length == 0) {
        getCurrentAdpData(Week, Year, _ApbAllitems, Count, filterkeys);
        // } else {
        //   let ApbOnloadFilter = APBOnloadFilter([..._ApbAllitems], filterkeys);
        //   setApbData([...ApbOnloadFilter]);
        //   sortApbDataArr = ApbOnloadFilter;
        //   setApbMasterData([...ApbOnloadFilter]);
        //   let ApbFilter = ActivityProductionBoardFilter(
        //     [...ApbOnloadFilter],
        //     filterkeys
        //   );
        //   reloadFilterOptions([...ApbFilter]);
        //   setApbFilterData(ApbFilter);
        //   sortApbFilterArr = ApbFilter;
        //   Activitypaginate(1, ApbFilter);
        //   setApbLoader(false);
        // }
      })
      .catch((error) => {
        ApbErrorFunction(error, "getCurrentApbData");
      });
  };
  const getCurrentAdpData = (Week, Year, data, Count, filterkeys) => {
    sharepointWeb.lists
      .getByTitle("Activity Delivery Plan")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
      )
      .expand("Developer,FieldValuesAsText")
      .filter("ActivityPlanID eq '" + Apb_ActivityPlanId + "' ")
      .top(5000)
      .get()
      .then((items) => {
        let _ApbAllitems = data;
        let count = Count;
        console.log(items, "Activity Delivery Plan");
        items.forEach(async (item, Index) => {
          if (
            _ApbAllitems.findIndex((pb) => {
              return pb.ActivityDeliveryPlanID == item.ID;
            }) == -1
          ) {
            let curATPDetails = gblATPDetails.filter((arr) => {
              return arr.ID == item.ActivityPlanID;
            });

            // const inputFormat44 = "YYYY-MM-DDTHH:mm:ss[Z]";
            // const outputFormat44 = "DD/MM/YYYY";


            // const reponseDate = await getCamelquery(item.ID);

            // let reponseDateActuall = moment(
            //   item["FieldValuesAsText"].EndDate,
            //   DateListFormat
            // ).format(DateListFormat);


            // if (reponseDate) {
            //   reponseDateActuall = reponseDate
            // }

            _ApbAllitems.push({
              RefId: count++,
              ID: 0,
              StartDate: moment(
                item["FieldValuesAsText"].StartDate,
                DateListFormat
              ).format(DateListFormat),
              EndDate: moment(
                item["FieldValuesAsText"].EndDate,
                DateListFormat
              ).format(DateListFormat),
              // StartDate: moment(
              //   item["FieldValuesAsText"].StartDate,
              //   DateListFormat
              // ).format(DateListFormat),
              // EndDate: `12/12/2025`,
              Source: "AP",
              Project:
                curATPDetails.length > 0
                  ? curATPDetails[0].Project
                  : item.Project
                    ? item.Project +
                    " " +
                    (item.ProjectVersion ? item.ProjectVersion : "V1")
                    : "",
              Product:
                curATPDetails.length > 0
                  ? curATPDetails[0].Product
                  : item.Product
                    ? item.Product +
                    " " +
                    (item.ProductVersion ? item.ProductVersion : "V1")
                    : "",
              Title: item.Types,
              PlannedHours: item.PlannedHours ? item.PlannedHours : 0,
              Monday: "0",
              Tuesday: "0",
              Wednesday: "0",
              Thursday: "0",
              Friday: "0",
              Saturday: "0",

              Sunday: "0",

              ActualHours: 0,
              DeveloperId: item.DeveloperId,
              DeveloperEmail: item.Developer ? item.Developer.EMail : "",
              Week: Week,
              Year: Year,
              Status: null,
              Lessons: item.Lesson,
              Steps: item.Title,
              ActivityPlanID: item.ActivityPlanID,
              ActivityDeliveryPlanID: item.ID,
              ADPActualHours: item.ActualHours ? item.ActualHours : 0,
              UnPlannedHours: item.UnPlannedHours ? item.UnPlannedHours : false,
              PHWeek: item.PHWeek ? item.PHWeek : null,
              Onchange: false,
            });
          }
        });
        let ApbOnloadFilter = APBOnloadFilter([..._ApbAllitems], filterkeys);
        setApbData([...ApbOnloadFilter]);
        sortApbDataArr = ApbOnloadFilter;
        setApbMasterData([...ApbOnloadFilter]);
        reloadFilterOptions([...ApbOnloadFilter]);
        let ApbFilter = ActivityProductionBoardFilter(
          [...ApbOnloadFilter],
          filterkeys
        );
        setApbFilterData(ApbFilter);
        sortApbFilterArr = ApbFilter;
        Activitypaginate(1, ApbFilter);
        setApbLoader(false);
      })
      .catch((error) => {
        ApbErrorFunction(error, "getCurrentAdpData");
      });
  };
  const getApbData = (devID, Week, Year, filterkeys) => {
    setApbLoader(true);
    sharepointWeb.lists
      .getByTitle("ActivityProductionBoard")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
      )
      .expand("Developer,FieldValuesAsText")
      // .filter(
      //   "Week eq '" +
      //     Apb_WeekNumber +
      //     "' and Year eq '" +
      //     Apb_Year +
      //     "' and Developer/EMail eq '" +
      //     loggeduseremail +
      //     "' "
      // )
      // .filter(
      //   "Week eq '" + Apb_WeekNumber + "' and Year eq '" + Apb_Year + "' "
      // )
      .filter(`SPFxFilter eq '${devID}-${Week}-${Year}'`)
      .top(5000)
      .get()
      .then((items) => {
        console.log(items, "ActivityProductionBoard");
        let _ApbAllitems = [];
        let Count = 0;
        items.forEach(async (item, Index) => {
          let curATPDetails = gblATPDetails.filter((arr) => {
            return arr.ID == item.ActivityPlanID;
          });

          // For onchange calculation
          let oldProduct =
            item.Product +
            " " +
            (item.ProductVersion ? item.ProductVersion : "V1");
          let oldProject =
            item.Project +
            " " +
            (item.ProjectVersion ? item.ProjectVersion : "V1");
          let NewProject =
            curATPDetails.length > 0 ? curATPDetails[0].Project : "";
          let NewProduct =
            curATPDetails.length > 0 ? curATPDetails[0].Product : "";

          if (
            //for Deleted ActivityPlan
            (curATPDetails.length > 0 && item.ActivityPlanID) ||
            (item.Project && !item.ActivityPlanID)
          ) {

            // const reponseDate = await getCamelquery(item.ActivityDeliveryPlanID);
            // let reponseDateActuall = moment(
            //   item["FieldValuesAsText"].EndDate,
            //   DateListFormat
            // ).format(DateListFormat);


            // if (reponseDate) {
            //   reponseDateActuall = reponseDate
            // }


            _ApbAllitems.push({
              RefId: Index + 1,
              ID: item.ID,
              StartDate: moment(
                item["FieldValuesAsText"].StartDate,
                DateListFormat
              ).format(DateListFormat),
              EndDate: moment(
                item["FieldValuesAsText"].EndDate,
                DateListFormat
              ).format(DateListFormat),
              Source: item.Source,
              Project:
                curATPDetails.length > 0
                  ? curATPDetails[0].Project
                  : item.Project
                    ? item.Project +
                    " " +
                    (item.ProjectVersion ? item.ProjectVersion : "V1")
                    : "",
              Product:
                curATPDetails.length > 0
                  ? curATPDetails[0].Product
                  : item.Product
                    ? item.Product +
                    " " +
                    (item.ProductVersion ? item.ProductVersion : "V1")
                    : "",
              Title: item.Title,
              PlannedHours: item.PlannedHours ? item.PlannedHours : 0,
              Monday: item.Monday ? item.Monday : "0",
              Tuesday: item.Tuesday ? item.Tuesday : "0",
              Wednesday: item.Wednesday ? item.Wednesday : "0",
              Thursday: item.Thursday ? item.Thursday : "0",
              Friday: item.Friday ? item.Friday : "0",
              Saturday: item.Saturday ? item.Saturday : "0",

              Sunday: item.Sunday ? item.Sunday : "0",

              ActualHours: item.ActualHours,
              DeveloperId: item.DeveloperId,
              DeveloperEmail: item.Developer ? item.Developer.EMail : "",
              Week: item.Week,
              Year: item.Year,
              Status: item.Status,
              Lessons: item.Lessons,
              Steps: item.Steps,
              ActivityPlanID: item.ActivityPlanID,
              ActivityDeliveryPlanID: item.ActivityDeliveryPlanID,
              ADPActualHours: item.ADPActualHours ? item.ADPActualHours : 0,
              UnPlannedHours: item.UnPlannedHours ? item.UnPlannedHours : false,
              PHWeek: item.PHWeek ? item.PHWeek : null,
              Onchange:
                oldProduct != NewProduct || oldProject != NewProject
                  ? true
                  : false,
            });
          }
        });
        getAdpData(Week, Year, _ApbAllitems, Count, devID, filterkeys);
      })
      .catch((error) => {
        ApbErrorFunction(error, "getApbData");
      });
  };


  // const getCamelquery = async (_id) => {



  //   let camelQueryXML: string =
  //     '<View>' +
  //     "<ViewFields>" +
  //     "<FieldRef Name='ID'/>" +
  //     "<FieldRef Name='auditResponseDate'/>" +
  //     "</ViewFields>" +
  //     `<Query>
  //     <Where>
  //        <And>
  //           <Eq>
  //              <FieldRef Name='DeliveryPlanID' />
  //              <Value Type='Number'>${_id}</Value>
  //           </Eq>
  //           <Or>
  //              <Eq>
  //                 <FieldRef Name='auditResponseType' />
  //                 <Value Type='Choice'>Signed Off</Value>
  //              </Eq>
  //              <Eq>
  //                 <FieldRef Name='auditResponseType' />
  //                 <Value Type='Choice'>Publish ready</Value>
  //              </Eq>
  //           </Or>
  //        </And>
  //     </Where>
  //  </Query>` +
  //     '</View>';


  //   //sp.web.lists.getByTitle("ProductionBoard").getItemsByCAMLQuery({ 'ViewXml': camelQueryXML }).then((productionBoardResponse: any)

  //   await sharepointWeb.lists
  //     .getByTitle("Review log").getItemsByCAMLQuery({ 'ViewXml': camelQueryXML }).then((data: any) => {
  //       if (data.length > 0) {
  //         console.log(moment(data[0]?.auditResponseDate).format("DD/MM/YYYY"), "-------------------------");
  //         return moment(data[0]?.auditResponseDate).format("DD/MM/YYYY");


  //       }

  //     });
  //   return null

  // };



  const getAdpData = (Week, Year, data, Count, devID, filterkeys) => {
    sharepointWeb.lists
      .getByTitle("Activity Delivery Plan")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
      )
      .expand("Developer,FieldValuesAsText")
      // .filter("DeveloperId eq '" + loggeduserid + "' ")
      // .filter("'" + today + "'ge StartDate and '" + today + "' le EndDate")
      .filter(`SPFxFilter eq '${devID}-0'`)
      .top(5000)
      .get()
      .then(async (items) => {
        let _ApbAllitems = data;
        let count = Count;
        console.log(items, "Activity Delivery Plan");
        // let _ApbAllitems = [];
        // let count = 0;

        // let getActivityID = records.reduce(function (item, e1) {
        //   var matches = item.filter(function (e2) {
        //     return e1.ActivityPlanID === e2.ActivityPlanID;
        //   });
        //   if (matches.length == 0) {
        //     item.push(e1);
        //   }
        //   return item;
        // }, []);
        // if (getActivityID.length > 0) {
        //   await getActivityID.forEach(async (items) => {
        //     await sharepointWeb.lists
        //       .getByTitle("Activity Delivery Plan")
        //       .items.select("*,Developer/Title,Developer/Id,Developer/EMail")
        //       .expand("Developer")
        //       .filter("ActivityPlanID eq '" + items.ActivityPlanID + "' ")
        //       .top(5000)
        //       .get()
        //       .then((items) => {
        items.forEach(async (item, Index) => {
          if (
            _ApbAllitems.findIndex((pb) => {
              return pb.ActivityDeliveryPlanID == item.ID;
            }) == -1
          ) {
            let curATPDetails = gblATPDetails.filter((arr) => {
              return arr.ID == item.ActivityPlanID;
            });

            if (
              //for Deleted ActivityPlan
              (curATPDetails.length > 0 && item.ActivityPlanID) ||
              (item.Project && !item.ActivityPlanID)
            ) {


              // const inputFormat44 = "YYYY-MM-DDTHH:mm:ss[Z]";
              // const outputFormat44 = "DD/MM/YYYY";


              // const reponseDate = await getCamelquery(item.ID);

              // let reponseDateActuall = ;


              // if (reponseDate) {
              //   reponseDateActuall = reponseDate
              // }


              _ApbAllitems.push({
                RefId: count++,
                ID: 0,
                // StartDate: moment(
                //   item["FieldValuesAsText"].StartDate,
                //   DateListFormat
                // ).format(DateListFormat),
                // StartDate: moment(
                //   item["FieldValuesAsText"].StartDate,
                //   DateListFormat
                // ).format(DateListFormat),
                StartDate: moment(
                  item["FieldValuesAsText"].StartDate,
                  DateListFormat
                ).format(DateListFormat),
                EndDate: moment(
                  item["FieldValuesAsText"].EndDate,
                  DateListFormat
                ).format(DateListFormat),
                Source: item.ActivityPlanID ? "AP" : "Ad hoc",
                Project:
                  curATPDetails.length > 0
                    ? curATPDetails[0].Project
                    : item.Project
                      ? item.Project +
                      " " +
                      (item.ProjectVersion ? item.ProjectVersion : "V1")
                      : "",
                Product:
                  curATPDetails.length > 0
                    ? curATPDetails[0].Product
                    : item.Product
                      ? item.Product +
                      " " +
                      (item.ProductVersion ? item.ProductVersion : "V1")
                      : "",
                Title: item.Types,
                PlannedHours: item.PlannedHours ? item.PlannedHours : 0,
                Monday: "0",
                Tuesday: "0",
                Wednesday: "0",
                Thursday: "0",
                Friday: "0",
                Saturday: "0",

                Sunday: "0",

                ActualHours: 0,
                DeveloperId: item.DeveloperId,
                DeveloperEmail: item.Developer ? item.Developer.EMail : "",
                Week: Week,
                Year: Year,
                Status: null,
                Lessons: item.Lesson,
                Steps: item.Title,
                ActivityPlanID: item.ActivityPlanID,
                ActivityDeliveryPlanID: item.ID,
                ADPActualHours: item.ActualHours ? item.ActualHours : 0,
                UnPlannedHours: item.UnPlannedHours
                  ? item.UnPlannedHours
                  : false,
                PHWeek: item.PHWeek ? item.PHWeek : null,
                Onchange: false,
              });
            }
          }
        });
        let ApbOnloadFilter = APBOnloadFilter([..._ApbAllitems], filterkeys);
        setApbData([...ApbOnloadFilter]);
        sortApbDataArr = ApbOnloadFilter;
        setApbMasterData([...ApbOnloadFilter]);
        reloadFilterOptions([...ApbOnloadFilter]);
        let ApbFilter = ActivityProductionBoardFilter(
          [...ApbOnloadFilter],
          filterkeys
        );
        setApbFilterData(ApbFilter);
        sortApbFilterArr = ApbFilter;
        Activitypaginate(1, ApbFilter);
        setApbLoader(false);
      });
    //     });
    //   } else {
    //     setApbLoader(false);
    //   }
    // })
    // .catch((error) => {
    //   ApbErrorFunction(error, "getModalBoxOptions1");
    // });
  };
  const saveApbData = () => {
    setApbLoader(true);
    let successCount = 0;
    ApbData.forEach((Apb, Index: number) => {
      let strDWYNA: string = `${Apb.DeveloperId}-${Apb.Week}-${Apb.Year}`;

      // Versions
      let PrjData = ProjectOrProductDetails.filter((arr) => {
        return (arr.Type = "Project" && arr.Key == Apb.Project);
      });
      let PrdData = ProjectOrProductDetails.filter((arr) => {
        return (arr.Type = "Product" && arr.Key == Apb.Product);
      });

      let PrjTitle =
        PrjData.length > 0 ? PrjData[0].Title : Apb.Project.replace("V1", "");
      let PrjVersion = PrjData.length > 0 ? PrjData[0].Version : "V1";

      let PrdTitle =
        PrdData.length > 0 ? PrdData[0].Title : Apb.Product.replace("V1", "");
      let PrdVersion = PrdData.length > 0 ? PrdData[0].Version : "V1";

      let requestdata = {
        StartDate: Apb.StartDate
          ? moment(Apb.StartDate, DateListFormat).format("YYYY-MM-DD")
          : null,
        EndDate: Apb.EndDate
          ? moment(Apb.EndDate, DateListFormat).format("YYYY-MM-DD")
          : null,
        Source: Apb.Source ? Apb.Source : null,
        Product: PrdTitle ? PrdTitle : null,
        Project: PrjTitle ? PrjTitle : null,
        ProductVersion: PrdVersion,
        ProjectVersion: PrjVersion,
        Title: Apb.Title ? Apb.Title : null,
        PlannedHours: Apb.PlannedHours ? Apb.PlannedHours : null,
        Monday: Apb.Monday ? Apb.Monday : "0",
        Tuesday: Apb.Tuesday ? Apb.Tuesday : "0",
        Wednesday: Apb.Wednesday ? Apb.Wednesday : "0",
        Thursday: Apb.Thursday ? Apb.Thursday : "0",
        Friday: Apb.Friday ? Apb.Friday : "0",
        Saturday: Apb.Saturday ? Apb.Saturday : "0",
        Sunday: Apb.Sunday ? Apb.Sunday : "0",
        ActualHours: Apb.ActualHours ? Apb.ActualHours : 0,
        DeveloperId: Apb.DeveloperId ? Apb.DeveloperId : null,
        Week: Apb.Week,
        Year: Apb.Year,
        Status: Apb.Status ? Apb.Status : null,
        Lessons: Apb.Lessons ? Apb.Lessons : null,
        Steps: Apb.Steps ? Apb.Steps : null,
        ActivityPlanID: Apb.ActivityPlanID ? Apb.ActivityPlanID : null,
        ActivityDeliveryPlanID: Apb.ActivityDeliveryPlanID
          ? Apb.ActivityDeliveryPlanID
          : null,
        ADPActualHours: Apb.ADPActualHours ? Apb.ADPActualHours : 0,
        UnPlannedHours: Apb.UnPlannedHours ? Apb.UnPlannedHours : false,
        PHWeek: Apb.PHWeek ? Apb.PHWeek : null,
        SPFxFilter: strDWYNA,
      };
      let AH =
        parseFloat(Apb.ADPActualHours ? Apb.ADPActualHours : 0) +
        parseFloat(Apb.ActualHours ? Apb.ActualHours : 0);

      // Adhoc task
      let strDSNA: string = `${Apb.DeveloperId}-0`;
      let responseDataforAPB;
      if (Apb.Source == "Ad hoc") {
        responseDataforAPB = {
          ActivityPlanID: "",
          Types: Apb.Title ? Apb.Title : "",
          PlannedHours: Apb.PlannedHours ? Apb.PlannedHours : 0,
          MinPH: 0,
          MaxPH: 0,
          Product: PrdTitle ? PrdTitle : null,
          Project: PrjTitle ? PrjTitle : null,
          ProductVersion: PrdVersion,
          ProjectVersion: PrjVersion,
          Lesson: Apb.Lessons ? Apb.Lessons : "",
          StartDate: Apb.StartDate
            ? moment(Apb.StartDate, DateListFormat).format("YYYY-MM-DD")
            : moment().format("YYYY-MM-DD"),
          EndDate: Apb.EndDate
            ? moment(Apb.EndDate, DateListFormat).format("YYYY-MM-DD")
            : moment().format("YYYY-MM-DD"),
          DeveloperId: Apb.DeveloperId ? Apb.DeveloperId : null,
          Status: "Scheduled",
          ActualHours: AH,
          OrderId: 0,
          LessonID: 0,
          Title: Apb.Steps ? Apb.Steps : "",
          SPFxFilter: strDSNA,
          UnPlannedHours: Apb.UnPlannedHours ? Apb.UnPlannedHours : false,
        };
      } else {
        responseDataforAPB = {
          ActualHours: AH,
        };
      }

      console.log(Apb.ActualHours);
      if (Apb.ID != 0 && Apb.Onchange == true) {
        sharepointWeb.lists
          .getByTitle("ActivityProductionBoard")
          .items.getById(Apb.ID)
          .update(requestdata)
          .then(() => {
            if (Apb.ActivityDeliveryPlanID) {
              sharepointWeb.lists
                .getByTitle("Activity Delivery Plan")
                .items.getById(Apb.ActivityDeliveryPlanID)
                .update(responseDataforAPB)
                .then((e) => { })
                .catch((error) => {
                  ApbErrorFunction(error, "saveApbData1");
                });
            }

            successCount++;
            if (successCount == ApbData.length) {
              setApbLoader(false);
              setApbMasterData([...ApbData]);
              let ApbFilter = ActivityProductionBoardFilter(
                [...ApbData],
                ApbFilterKeys
              );
              setApbFilterData(ApbFilter);
              sortApbFilterArr = ApbFilter;
              Activitypaginate(1, ApbFilter);
              // setApbFilterOptions({ ...ApbFilterKeys });
              setApbUpdate(!ApbUpdate);
              sortApbUpdate = !ApbUpdate;
              AddSuccessPopup();
            }
          })
          .catch((error) => {
            ApbErrorFunction(error, "saveApbData2");
          });
      } else if (Apb.ID == 0) {
        if (Apb.Source == "Ad hoc" && Apb.ActivityDeliveryPlanID == null) {
          sharepointWeb.lists
            .getByTitle("Activity Delivery Plan")
            .items.add(responseDataforAPB)
            .then((ev) => {
              requestdata.ActivityDeliveryPlanID = ev.data.Id;
              ApbData[Index].ActivityDeliveryPlanID = ev.data.Id;

              sharepointWeb.lists
                .getByTitle("ActivityProductionBoard")
                .items.add(requestdata)
                .then((e) => {
                  if (Apb.ActivityDeliveryPlanID) {
                    sharepointWeb.lists
                      .getByTitle("Activity Delivery Plan")
                      .items.getById(Apb.ActivityDeliveryPlanID)
                      .update({ ActualHours: AH })
                      .then((e) => { })
                      .catch((error) => {
                        ApbErrorFunction(error, "saveApbData3");
                      });
                  }
                  successCount++;
                  ApbData[Index].ID = e.data.ID;
                  if (successCount == ApbData.length) {
                    setApbLoader(false);
                    setApbData([...ApbData]);
                    sortApbDataArr = ApbData;
                    setApbMasterData([...ApbData]);
                    let ApbFilter = ActivityProductionBoardFilter(
                      [...ApbData],
                      ApbFilterKeys
                    );
                    setApbFilterData(ApbFilter);
                    sortApbFilterArr = ApbFilter;
                    Activitypaginate(1, ApbFilter);
                    // setApbFilterOptions({ ...ApbFilterKeys });
                    setApbUpdate(!ApbUpdate);
                    sortApbUpdate = !ApbUpdate;
                    AddSuccessPopup();
                  }
                })
                .catch((error) => {
                  ApbErrorFunction(error, "saveApbData4");
                });
            });
        } else {
          sharepointWeb.lists
            .getByTitle("ActivityProductionBoard")
            .items.add(requestdata)
            .then((e) => {
              if (Apb.ActivityDeliveryPlanID) {
                sharepointWeb.lists
                  .getByTitle("Activity Delivery Plan")
                  .items.getById(Apb.ActivityDeliveryPlanID)
                  .update({ ActualHours: AH })
                  .then((e) => { })
                  .catch((error) => {
                    ApbErrorFunction(error, "saveApbData3");
                  });
              }
              successCount++;
              ApbData[Index].ID = e.data.ID;
              if (successCount == ApbData.length) {
                setApbLoader(false);
                setApbData([...ApbData]);
                sortApbDataArr = ApbData;
                setApbMasterData([...ApbData]);
                let ApbFilter = ActivityProductionBoardFilter(
                  [...ApbData],
                  ApbFilterKeys
                );
                setApbFilterData(ApbFilter);
                sortApbFilterArr = ApbFilter;
                Activitypaginate(1, ApbFilter);
                // setApbFilterOptions({ ...ApbFilterKeys });
                setApbUpdate(!ApbUpdate);
                sortApbUpdate = !ApbUpdate;
                AddSuccessPopup();
              }
            })
            .catch((error) => {
              ApbErrorFunction(error, "saveApbData4");
            });
        }
      } else {
        successCount++;
        if (successCount == ApbData.length) {
          setApbLoader(false);
          setApbUpdate(!ApbUpdate);
          sortApbUpdate = !ApbUpdate;
          AddSuccessPopup();
        }
      }
    });
  };
  const saveApbDRData = () => {
    let requestdata = {
      Link: ApbDocumentReview.Link,
      Request: ApbDocumentReview.Request ? ApbDocumentReview.Request : null,
      RequesttoId: ApbDocumentReview.Requestto
        ? ApbDocumentReview.Requestto
        : null,
      EmailccId: ApbDocumentReview.Emailcc
        ? { results: ApbDocumentReview.Emailcc }
        : { results: [] },
      Project: ApbDocumentReview.Project ? ApbDocumentReview.Project : null,
      Documenttype: ApbDocumentReview.Documenttype
        ? ApbDocumentReview.Documenttype
        : null,
      Comments: ApbDocumentReview.Comments ? ApbDocumentReview.Comments : null,
      Confidential: ApbDocumentReview.Confidential,
      IsExternalAllow: ApbDocumentReview.IsExternalAllow,
      Product: ApbDocumentReview.Product ? ApbDocumentReview.Product : null,
      AnnualPlanID: ApbDocumentReview.AnnualPlanID
        ? ApbDocumentReview.AnnualPlanID
        : null,
      DeliveryPlanID: ApbDocumentReview.DeliveryPlanID
        ? ApbDocumentReview.DeliveryPlanID
        : null,
      ProductionBoardID: ApbDocumentReview.ProductionBoardID
        ? ApbDocumentReview.ProductionBoardID
        : null,
      DRPageName: "Activity Plan",
    };
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .items.add(requestdata)
      .then((e) => {
        if (ApbDocumentReview.ProductionBoardID) {
          sharepointWeb.lists
            .getByTitle("ActivityProductionBoard")
            .items.getById(ApbDocumentReview.ProductionBoardID)
            .update({ Status: "Pending" })
            .then(() => {
              let Index = ApbData.findIndex(
                (obj) => obj.ID == ApbDocumentReview.ProductionBoardID
              );
              ApbData[Index].Status = "Pending";
              setApbData([...ApbData]);
              sortApbDataArr = ApbData;
              setDocumentLinkStatus("no-checked")
            })
            .catch((error) => {
              setDocumentLinkStatus("no-checked")
              ApbErrorFunction(error, "saveApbDRData1");
            });
        }
        setApbModalBoxVisibility(false);
        AddDRSuccessPopup();
      })
      .catch((error) => {
        ApbErrorFunction(error, "saveApbDRData2");
      });
  };
  const cancelApbData = () => {
    setDocumentLinkStatus("no-checked")
    // setApbFilterOptions({ ...ApbFilterKeys });
    setApbData([...ApbMasterData]);
    sortApbDataArr = ApbMasterData;
    setApbUpdate(false);
    sortApbUpdate = false;
    let ApbFilter = ActivityProductionBoardFilter(
      [...ApbMasterData],
      ApbFilterKeys
    );
    setApbFilterData(ApbFilter);
    sortApbFilterArr = ApbFilter;
    Activitypaginate(1, ApbFilter);
    setApbAutoSave(false);
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;
    tempArrReload.forEach((item, Index) => {
      if (Apb_ActivityPlanId && Index == 0) {
        FilterProject = item.Project;
        FilterProduct = item.Product;
      }

      if (
        ApbDrpDwnOptns.Lessons.findIndex((BA) => {
          return BA.key == item.Lessons;
        }) == -1 &&
        item.Lessons
      ) {
        ApbDrpDwnOptns.Lessons.push({
          key: item.Lessons,
          text: item.Lessons,
        });
      }
      if (
        ApbDrpDwnOptns.Steps.findIndex((Source) => {
          return Source.key == item.Steps;
        }) == -1 &&
        item.Steps
      ) {
        ApbDrpDwnOptns.Steps.push({
          key: item.Steps,
          text: item.Steps,
        });
      }
      if (
        ApbDrpDwnOptns.Product.findIndex((Product) => {
          return Product.key == item.Product;
        }) == -1 &&
        item.Product
      ) {
        ApbDrpDwnOptns.Product.push({
          key: item.Product,
          text: item.Product,
        });
      }
      if (
        ApbDrpDwnOptns.Project.findIndex((Project) => {
          return Project.key == item.Project;
        }) == -1 &&
        item.Project
      ) {
        ApbDrpDwnOptns.Project.push({
          key: item.Project,
          text: item.Project,
        });
      }
      if (Apb_ActivityPlanId) {
        if (
          ApbDrpDwnOptns.Developer.findIndex((Developer) => {
            return Developer.key == item.DeveloperId;
          }) == -1 &&
          item.DeveloperId &&
          item.DeveloperEmail != "lally@goodtogreatschools.org.au"
        ) {
          let devName = allPeoples.filter(
            (dev) => dev.ID == item.DeveloperId
          )[0].text;
          ApbDrpDwnOptns.Developer.push({
            key: item.DeveloperId,
            text: devName,
          });
        }
      }
    });

    if (!Apb_ActivityPlanId) {
      allPeoples.forEach((arr) => {
        if (
          ApbDrpDwnOptns.Developer.findIndex((Developer) => {
            return Developer.key == arr.ID;
          }) == -1 &&
          arr.ID &&
          arr.secondaryText != "lally@goodtogreatschools.org.au" &&
          arr.secondaryText != ""
        ) {
          ApbDrpDwnOptns.Developer.push({
            key: arr.ID,
            text: arr.text,
          });
        }
      });
    }
    ApbDrpDwnOptns.Developer = usersOrderFunction(ApbDrpDwnOptns.Developer);

    const _sortFilterKeys = (a, b) => {
      if (a.key < b.key) {
        return -1;
      }
      if (a.key > b.key) {
        return 1;
      }
      return 0;
    };

    let maxWeek = 53;
    for (let i = 1; i <= maxWeek; i++) {
      if (
        ApbDrpDwnOptns.WeekNumber.findIndex((arr) => {
          return arr.key == i;
        }) == -1
      ) {
        ApbDrpDwnOptns.WeekNumber.push({
          key: i,
          text: i.toString(),
        });
      }
    }
    for (let i = 2020; i < Apb_Year; i++) {
      if (
        ApbDrpDwnOptns.Year.findIndex((arr) => {
          return arr.key == i;
        }) == -1
      ) {
        ApbDrpDwnOptns.Year.push({
          key: i,
          text: i.toString(),
        });
      }
    }
    ApbDrpDwnOptns.WeekNumber.sort(_sortFilterKeys);
    ApbDrpDwnOptns.Year.sort(_sortFilterKeys);

    setApbDropDownOptions(ApbDrpDwnOptns);
  };
  const usersOrderFunction = (dropDown): any => {
    let nonArchived = dropDown.filter((user) => {
      return !user.text.includes("Archive");
    });

    let archived = dropDown.filter((user) => {
      return user.text.includes("Archive");
    });

    return nonArchived.concat(archived);
  };
  const drPBValidationFunction = () => {
    let isError = false;

    let errorStatus = {
      Request: "",
      Requestto: "",
      Documenttype: "",
      Link: "",
    };

    if (!ApbDocumentReview.Request) {
      isError = true;
      errorStatus.Request = "Please select a value for request";
    }
    if (!ApbDocumentReview.Requestto) {
      isError = true;
      errorStatus.Requestto = "Please select a value for request to";
    }
    if (!ApbDocumentReview.Documenttype) {
      isError = true;
      errorStatus.Documenttype = "Please select a value for document type";
    }
    if (!ApbDocumentReview.Link) {
      isError = true;
      errorStatus.Link = "Please enter a value for link";
    }
    if (ApbDocumentReview.Link && !ApbDocumentReview.IsExternalAllow) {


      const respV = isLinkValid(ApbDocumentReview.Link)

      if (!respV) {
        isError = true;
        console.log("incorrect on consoele")
        setDocumentLinkStatus("incorrect")
        // return !hasAspx; // Invalid if link has "aspx"
      } else {

        setDocumentLinkStatus("correct")
      }
      // return hasSiteOrSharePoint; // Valid if link has "site" or "sharepoint"



    }

    //here it will be link validatuin
    if (!isError) {
      setApbButtonLoader(true);
      saveApbDRData();
    } else {
      setAdrPBShowMessage(errorStatus);
    }
  };
  const ApbValidationFunction = () => {
    let tempArronchange = ApbAdhocPopup.value;
    let isError = false;

    let errorStatus = {
      Type: "",
      StartDate: "",
      EndDate: "",
      Project: "",
      Product: "",
      Lessons: "",
      Steps: "",
      PlannedHours: "",
    };

    if (!tempArronchange["Title"]) {
      isError = true;
      errorStatus.Type = "Please select a value for type";
    }
    if (!tempArronchange["StartDate"]) {
      isError = true;
      errorStatus.StartDate = "Please select a value for startDate";
    }
    if (!tempArronchange["EndDate"]) {
      isError = true;
      errorStatus.EndDate = "Please select a value for endDate";
    }
    if (!tempArronchange["Product"]) {
      isError = true;
      errorStatus.Product = "Please select a value for product";
    }
    if (!tempArronchange["Project"]) {
      isError = true;
      errorStatus.Project = "Please select a value for name of the deliverable";
    }
    if (!tempArronchange["Lessons"]) {
      isError = true;
      errorStatus.Lessons = "Please enter a value for session";
    }
    if (!tempArronchange["Steps"]) {
      isError = true;
      errorStatus.Steps = "Please enter a value for task";
    }
    if (!tempArronchange["PlannedHours"]) {
      isError = true;
      errorStatus.PlannedHours = "Please enter a value for planned hours";
    }

    if (!isError) {
      tempArronchange["StartDate"] = moment(
        tempArronchange["StartDate"]
      ).format(DateListFormat);
      tempArronchange["EndDate"] = moment(tempArronchange["EndDate"]).format(
        DateListFormat
      );
      if (ApbAdhocPopup.isNew) {
        setApbButtonLoader(true);
        setApbData(ApbData.concat(tempArronchange));
        reloadFilterOptions(ApbData.concat(tempArronchange));
        let pbFilter = ActivityProductionBoardFilter(
          [...ApbData.concat(tempArronchange)],
          ApbFilterOptions
        );
        setApbFilterData([...pbFilter]);
        Activitypaginate(1, pbFilter);
        setApbUpdate(true);
        setApbAdhocPopup({
          visible: false,
          isNew: ApbAdhocPopup.isNew,
          value: {},
        });

        //Sorting
        sortApbUpdate = true;
        sortApbFilterArr = pbFilter;
        sortApbDataArr = ApbData.concat(tempArronchange);
        setapbColumns(_apbColumns);
        setApbButtonLoader(false);
      } else {
        let Index = ApbData.findIndex(
          (obj) => obj.RefId == tempArronchange["RefId"]
        );
        ApbData[Index] = tempArronchange;

        setApbButtonLoader(true);
        setApbData([...ApbData]);
        reloadFilterOptions([...ApbData]);
        let pbFilter = ActivityProductionBoardFilter(
          [...ApbData],
          ApbFilterOptions
        );

        setApbFilterData([...pbFilter]);
        Activitypaginate(1, pbFilter);
        setApbUpdate(true);
        setApbAdhocPopup({
          visible: false,
          isNew: ApbAdhocPopup.isNew,
          value: {},
        });

        //Sorting
        sortApbUpdate = true;
        sortApbFilterArr = pbFilter;
        sortApbDataArr = ApbData;
        setapbColumns(_apbColumns);
        setApbButtonLoader(false);
      }
    } else {
      setApbShowMessage(errorStatus);
    }
  };


  function isLinkValid(link) {
    var hasSiteOrSharePoint = /site|sharepoint/i.test(link);
    if (hasSiteOrSharePoint) {
      var hasAspx = /aspx/i.test(link);
      return !hasAspx; // Invalid if link has "aspx"
    }
    return hasSiteOrSharePoint; // Valid if link has "site" or "sharepoint"
  }
  const ApbDeleteItem = (id: number) => {
    sharepointWeb.lists
      .getByTitle("ActivityProductionBoard")
      .items.getById(id)
      .delete()
      .then(() => {
        let AdpData = ApbMasterData.filter((arr) => {
          return arr.ID == id;
        });

        if (AdpData.length > 0) {
          sharepointWeb.lists
            .getByTitle("Activity Delivery Plan")
            .items.getById(AdpData[0].ActivityDeliveryPlanID)
            .delete()
            .then(() => {
              let tempMasterArr = [...ApbMasterData];
              let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
              tempMasterArr.splice(targetIndex, 1);

              let temp_ap_arr = [...ApbData];
              let targetIndexapdata = temp_ap_arr.findIndex(
                (arr) => arr.ID == id
              );
              temp_ap_arr.splice(targetIndexapdata, 1);

              setApbData([...temp_ap_arr]);
              sortApbDataArr = temp_ap_arr;
              setApbMasterData([...tempMasterArr]);
              reloadFilterOptions([...tempMasterArr]);
              let pbFilter = ActivityProductionBoardFilter(
                [...temp_ap_arr],
                ApbFilterOptions
              );

              setApbFilterData(pbFilter);
              sortApbFilterArr = pbFilter;
              Activitypaginate(1, pbFilter);
              setApbDeletePopup({ condition: false, targetId: 0 });
              DeleteSuccessPopup();
            })
            .catch((error) => {
              ApbErrorFunction(error, "ApbDeleteItem");
            });
        } else {
          let tempMasterArr = [...ApbMasterData];
          let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
          tempMasterArr.splice(targetIndex, 1);

          let temp_ap_arr = [...ApbData];
          let targetIndexapdata = temp_ap_arr.findIndex((arr) => arr.ID == id);
          temp_ap_arr.splice(targetIndexapdata, 1);

          setApbData([...temp_ap_arr]);
          sortApbDataArr = temp_ap_arr;
          setApbMasterData([...tempMasterArr]);
          reloadFilterOptions([...tempMasterArr]);
          let pbFilter = ActivityProductionBoardFilter(
            [...temp_ap_arr],
            ApbFilterOptions
          );

          setApbFilterData(pbFilter);
          sortApbFilterArr = pbFilter;
          Activitypaginate(1, pbFilter);
          setApbDeletePopup({ condition: false, targetId: 0 });
          DeleteSuccessPopup();
        }
      })
      .catch((error) => {
        ApbErrorFunction(error, "ApbDeleteItem");
      });
  };
  const ApbErrorFunction = (error, functionName) => {
    console.log(error);

    let response = {
      ComponentName: "Activity production board",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setApbLoader(false);
        setApbButtonLoader(false);
        setApbUpdate(!ApbUpdate);
        sortApbUpdate = !ApbUpdate;
        ErrorPopup();
      }
    );
  };
  const AddSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Production board is successfully submitted !!!")
  );
  const AddDRSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Document is successfully submitted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const DeleteSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Production board is successfully deleted !!!")
  );

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _apbColumns;
    const newColumns: IColumn[] = tempapColumns.slice();
    const currColumn: IColumn = newColumns.filter(
      (currCol) => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });

    const newApbDataArr = _copyAndSort(
      sortApbDataArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newApbFilterArr = _copyAndSort(
      sortApbFilterArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setApbData([...newApbDataArr]);
    setApbFilterData([...newApbFilterArr]);
    Activitypaginate(1, newApbFilterArr);
  };
  function _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    let key = columnKey as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
      );
  }

  //Onchange Function
  const GetUserDetails = (filterText, currentPersonas) => {
    let _allPeoples = allPeoples;

    _allPeoples = _allPeoples.filter((arr) => {
      return arr.text.toLowerCase().indexOf("archive") == -1;
    });

    if (currentPersonas.length > 0) {
      _allPeoples = _allPeoples.filter(
        (arr) => !currentPersonas.some((persona) => persona.ID == arr.ID)
      );
    }

    var result = _allPeoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };
  const doesTextStartWith = (text: string, filterText: string) => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };
  const onChangeFilter = (key, option) => {
    // let week;
    // let year;
    let tempArr = [...ApbData];
    let tempDpFilterKeys = { ...ApbFilterOptions };
    tempDpFilterKeys[`${key}`] = option;

    key == "WeekNumber" ? setApbWeek(option) : null;
    key == "Year" ? setApbYear(option) : null;

    // if (tempDpFilterKeys.Week == "This Week") {
    //   week = Apb_WeekNumber;
    //   year = Apb_Year;
    //   setApbWeek(Apb_WeekNumber);
    //   setApbYear(Apb_Year);
    // } else if (tempDpFilterKeys.Week == "Last Week") {
    //   week = Apb_LastWeekNumber;
    //   year = Apb_LastWeekYear;
    //   setApbWeek(Apb_LastWeekNumber);
    //   setApbYear(Apb_LastWeekYear);
    // } else if (tempDpFilterKeys.Week == "Next Week") {
    //   week = Apb_NextWeekNumber;
    //   year = Apb_NextWeekYear;
    //   setApbWeek(Apb_NextWeekNumber);
    //   setApbYear(Apb_NextWeekYear);
    // }

    if (Apb_ActivityPlanId) {
      key == "WeekNumber" || key == "Year"
        ? getCurrentApbData(
          tempDpFilterKeys.WeekNumber,
          tempDpFilterKeys.Year,
          tempDpFilterKeys
        )
        : null;
    } else {
      if (
        tempDpFilterKeys.Showonly == "All" &&
        (key == "Developer" || key == "WeekNumber" || key == "Year")
      ) {
        setApbFilterOptions({ ...tempDpFilterKeys });
        getApbData(
          tempDpFilterKeys.Developer,
          tempDpFilterKeys.WeekNumber,
          tempDpFilterKeys.Year,
          tempDpFilterKeys
        );
      } else if (
        tempDpFilterKeys.Showonly == "Mine" &&
        (key == "Showonly" || key == "WeekNumber" || key == "Year")
      ) {
        tempDpFilterKeys["Developer"] = loggeduserid;
        setApbFilterOptions({ ...tempDpFilterKeys });
        getApbData(
          loggeduserid,
          tempDpFilterKeys.WeekNumber,
          tempDpFilterKeys.Year,
          tempDpFilterKeys
        );
      }
    }
    setApbFilterOptions({ ...tempDpFilterKeys });
    let ApbFilter = ActivityProductionBoardFilter(
      [...tempArr],
      tempDpFilterKeys
    );
    setApbFilterData(ApbFilter);
    sortApbFilterArr = ApbFilter;
    Activitypaginate(1, ApbFilter);
  };
  const ApbOnchangeItems = (RefId, key, value) => {
    let Index = ApbData.findIndex((obj) => obj.RefId == RefId);
    let filIndex = ApbFilterData.findIndex((obj) => obj.RefId == RefId);
    let disIndex = ApbDisplayData.findIndex((obj) => obj.RefId == RefId);
    let ApbBeforeData = ApbData[Index];

    let ApbOnchangeData = [
      {
        ID: ApbBeforeData.ID,
        StartDate: ApbBeforeData.StartDate,
        EndDate: ApbBeforeData.EndDate,
        Source: ApbBeforeData.Source,
        Project: ApbBeforeData.Project,
        Product: ApbBeforeData.Product,
        Title: ApbBeforeData.Title,
        PlannedHours: ApbBeforeData.PlannedHours,
        Monday: key == "Monday" ? value : ApbBeforeData.Monday,
        Tuesday: key == "Tuesday" ? value : ApbBeforeData.Tuesday,
        Wednesday: key == "Wednesday" ? value : ApbBeforeData.Wednesday,
        Thursday: key == "Thursday" ? value : ApbBeforeData.Thursday,
        Friday: key == "Friday" ? value : ApbBeforeData.Friday,
        Saturday: key == "Saturday" ? value : ApbBeforeData.Saturday,
        Sunday: key == "Sunday" ? value : ApbBeforeData.Sunday,
        ActualHours: ApbBeforeData.ActualHours,
        DeveloperId: ApbBeforeData.DeveloperId,
        DeveloperEmail: ApbBeforeData.DeveloperEmail,
        RefId: ApbBeforeData.RefId,
        Week: ApbBeforeData.Week,
        Year: ApbBeforeData.Year,
        Status: ApbBeforeData.Status,
        Lessons: ApbBeforeData.Lessons,
        Steps: ApbBeforeData.Steps,
        ActivityPlanID: ApbBeforeData.ActivityPlanID,
        ActivityDeliveryPlanID: ApbBeforeData.ActivityDeliveryPlanID,
        ADPActualHours: ApbBeforeData.ADPActualHours,
        UnPlannedHours: ApbBeforeData.UnPlannedHours,
        PHWeek: ApbBeforeData.PHWeek,
        Onchange: true,
      },
    ];
    ApbOnchangeData[0]["ActualHours"] =
      parseFloat(
        !isNaN(ApbOnchangeData[0]["Monday"]) && ApbOnchangeData[0]["Monday"]
          ? ApbOnchangeData[0]["Monday"]
          : 0
      ) +
      parseFloat(
        !isNaN(ApbOnchangeData[0]["Tuesday"]) && ApbOnchangeData[0]["Tuesday"]
          ? ApbOnchangeData[0]["Tuesday"]
          : 0
      ) +
      parseFloat(
        !isNaN(ApbOnchangeData[0]["Wednesday"]) &&
          ApbOnchangeData[0]["Wednesday"]
          ? ApbOnchangeData[0]["Wednesday"]
          : 0
      ) +
      parseFloat(
        !isNaN(ApbOnchangeData[0]["Thursday"]) && ApbOnchangeData[0]["Thursday"]
          ? ApbOnchangeData[0]["Thursday"]
          : 0
      ) +
      parseFloat(
        !isNaN(ApbOnchangeData[0]["Friday"]) && ApbOnchangeData[0]["Friday"]
          ? ApbOnchangeData[0]["Friday"]
          : 0
      )+
      parseFloat(
        !isNaN(ApbOnchangeData[0]["Saturday"]) && ApbOnchangeData[0]["Saturday"]
          ? ApbOnchangeData[0]["Saturday"]
          : 0
      )+
      parseFloat(
        !isNaN(ApbOnchangeData[0]["Sunday"]) && ApbOnchangeData[0]["Sunday"]
          ? ApbOnchangeData[0]["Sunday"]
          : 0
      );

    ApbData[Index] = ApbOnchangeData[0];
    ApbFilterData[filIndex] = ApbOnchangeData[0];
    ApbDisplayData[disIndex] = ApbOnchangeData[0];
    setApbData([...ApbData]);
    sortApbDataArr = ApbData;
    setApbFilterData([...ApbFilterData]);
    sortApbFilterArr = ApbFilterData;
    setApbDisplayData([...ApbDisplayData]);
  };
  const AdrPBAddOnchange = (key, value) => {
    let tempArronchange = ApbDocumentReview;
    if (key == "Request") tempArronchange.Request = value;
    else if (key == "Requestto") tempArronchange.Requestto = value;
    else if (key == "Emailcc") tempArronchange.Emailcc = value;
    else if (key == "Documenttype") tempArronchange.Documenttype = value;
    else if (key == "Link") tempArronchange.Link = value;
    else if (key == "Comments") tempArronchange.Comments = value;
    else if (key == "Confidential") tempArronchange.Confidential = value;
    else if (key == "IsExternalAllow") tempArronchange.IsExternalAllow = value;

    console.log(tempArronchange);
    setApbDocumentReview(tempArronchange);
  };
  const ApbAddOnchange = (key, value) => {
    let tempArronchange = ApbAdhocPopup.value;
    if (key == "Title") tempArronchange["Title"] = value;
    else if (key == "StartDate") tempArronchange["StartDate"] = value;
    else if (key == "EndDate") tempArronchange["EndDate"] = value;
    else if (key == "Product") {
      tempArronchange["Product"] = value;
    } else if (key == "Project") tempArronchange["Project"] = value;
    else if (key == "Lessons") tempArronchange["Lessons"] = value;
    else if (key == "Steps") tempArronchange["Steps"] = value;
    else if (key == "PlannedHours") tempArronchange["PlannedHours"] = value;
    else if (key == "UnPlannedHours") tempArronchange["UnPlannedHours"] = value;
    setApbAdhocPopup({
      visible: true,
      isNew: ApbAdhocPopup.isNew,
      value: tempArronchange,
    });
    console.log(tempArronchange);
  };

  const Activitypaginate = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let ActivitypaginatedItems = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setApbDisplayData(ActivitypaginatedItems);
      setApbCurrentPage(pagenumber);
    } else {
      setApbDisplayData([]);
      setApbCurrentPage(1);
    }
  };
  const APBOnloadFilter = (data, filterValue) => {
    let tempADpFilterKeys = { ...filterValue };
    let tempArr = data;

    if (tempADpFilterKeys.WeekNumber) {
      tempArr = tempArr.filter((arr) => {
        // let start = moment(arr.StartDate).isoWeek();
        // let end = moment(arr.EndDate).isoWeek();
        // let today = tempADpFilterKeys.WeekNumber;
        // return today >= start && today <= end;

        let start = moment(arr.StartDate, DateListFormat)
          .year()
          .toString()
          .concat(
            (
              "0" + moment(arr.StartDate, DateListFormat).isoWeek().toString()
            ).slice(-2)
          );
        let end = moment(arr.EndDate, DateListFormat)
          .year()
          .toString()
          .concat(
            (
              "0" + moment(arr.EndDate, DateListFormat).isoWeek().toString()
            ).slice(-2)
          );
        let today = tempADpFilterKeys.Year.toString().concat(
          ("0" + tempADpFilterKeys.WeekNumber.toString()).slice(-2)
        );
        //   .year()
        //   .toString()
        //   .concat(("0" + tempADpFilterKeys.WeekNumber.toString()).slice(-2));

        return (
          parseInt(today) >= parseInt(start) && parseInt(today) <= parseInt(end)
        );
      });
    }
    if (tempADpFilterKeys.Year) {
      tempArr = tempArr.filter((arr) => {
        return arr.Year == tempADpFilterKeys.Year;
      });
    }
    // if (tempADpFilterKeys.Week == "This Week") {
    //   tempArr = tempArr.filter((arr) => {
    //     let start = moment(arr.StartDate).isoWeek();
    //     let end = moment(arr.EndDate).isoWeek();
    //     let today = Apb_WeekNumber;
    //     return today >= start && today <= end;
    //   });
    // } else if (tempADpFilterKeys.Week == "Last Week") {
    //   tempArr = tempArr.filter((arr) => {
    //     let start = moment(arr.StartDate).isoWeek();
    //     let end = moment(arr.EndDate).isoWeek();
    //     let today = Apb_LastWeekNumber;
    //     return today >= start && today <= end;
    //   });
    // } else if (tempADpFilterKeys.Week == "Next Week") {
    //   tempArr = tempArr.filter((arr) => {
    //     let start = moment(arr.StartDate).isoWeek();
    //     let end = moment(arr.EndDate).isoWeek();
    //     let today = Apb_NextWeekNumber;
    //     return today >= start && today <= end;
    //   });
    // }

    tempArr.forEach((arr, index) => {
      let dpBeforeData = tempArr[index];
      let dpOnchangeData = [
        {
          RefId: dpBeforeData.RefId,
          ID: dpBeforeData.ID,
          StartDate: dpBeforeData.StartDate,
          EndDate: dpBeforeData.EndDate,
          Source: dpBeforeData.Source,
          Project: dpBeforeData.Project,
          Product: dpBeforeData.Product,
          Title: dpBeforeData.Title,
          PlannedHours: dpBeforeData.PlannedHours,
          Monday: dpBeforeData.Monday,
          Tuesday: dpBeforeData.Tuesday,
          Wednesday: dpBeforeData.Wednesday,
          Thursday: dpBeforeData.Thursday,
          Friday: dpBeforeData.Friday,
          Saturday: dpBeforeData.Saturday,

          Sunday: dpBeforeData.Sunday,

          ActualHours: dpBeforeData.ActualHours,
          DeveloperId: dpBeforeData.DeveloperId,
          DeveloperEmail: dpBeforeData.DeveloperEmail,
          Week: dpBeforeData.Week,
          Year: dpBeforeData.Year,
          Status: dpBeforeData.Status,
          Lessons: dpBeforeData.Lessons,
          Steps: dpBeforeData.Steps,
          ActivityPlanID: dpBeforeData.ActivityPlanID,
          ActivityDeliveryPlanID: dpBeforeData.ActivityDeliveryPlanID,
          ADPActualHours: dpBeforeData.ADPActualHours,
          UnPlannedHours: dpBeforeData.UnPlannedHours,
          PHWeek: dpBeforeData.PHWeek,
          Onchange: true,
        },
      ];
      tempArr[index] = dpOnchangeData[0];
    });

    return tempArr;
  };
  const ActivityProductionBoardFilter = (data, filterValue) => {
    let tempArr = data;
    let tempADpFilterKeys = { ...filterValue };

    if (tempADpFilterKeys.Showonly == "Mine") {
      tempArr = tempArr.filter((arr) => {
        return arr.DeveloperEmail == loggeduseremail;
      });
    }

    if (tempADpFilterKeys.Showonly == "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.DeveloperId == tempADpFilterKeys.Developer;
      });
    }
    if (tempADpFilterKeys.Lessons != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Lessons == tempADpFilterKeys.Lessons;
      });
    }
    if (tempADpFilterKeys.Steps != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Steps == tempADpFilterKeys.Steps;
      });
    }
    if (tempADpFilterKeys.Product != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Product == tempADpFilterKeys.Product;
      });
    }
    if (tempADpFilterKeys.Project != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Project == tempADpFilterKeys.Project;
      });
    }

    return tempArr;
  };
  const sumOfHours = () => {
    var sum: number = 0;
    // let tempArr = ApbFilterData;
    let tempArr = ApbFilterData.filter((arr) => {
      return arr.UnplannedHours != true;
    });
    if (tempArr.length > 0) {
      tempArr.forEach((x) => {
        sum += parseFloat(x.PlannedHours ? x.PlannedHours : 0);
      });
      return sum % 1 == 0 ? sum : sum.toFixed(2);
    } else {
      return 0;
    }
  };
  const sumOfActualHours = () => {
    var sum: number = 0;
    // let tempArr = ApbFilterData;
    let tempArr = ApbFilterData.filter((arr) => {
      return arr.UnplannedHours != true;
    });
    if (tempArr.length > 0) {
      tempArr.forEach((x) => {
        sum += parseFloat(x.ActualHours ? x.ActualHours : 0);
      });
      return sum % 1 == 0 ? sum : sum.toFixed(2);
    } else {
      return sum ? sum : 0;
    }
  };
  // Return function
  return (
    <>
      {ApbLoader ? (
        <CustomLoader />
      ) : (
        <div style={{ padding: "5px 15px" }}>
          {/* {ApbLoader ? <CustomLoader /> : null} */}
          <div
            className={styles.apHeaderSection}
            style={{ paddingBottom: "0" }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                color: "#2392b2",
              }}
            >
              <div className={styles.dpTitle}>
                {Apb_ActivityPlanId ? (
                  <Icon
                    aria-label="ChevronLeftMed"
                    iconName="NavigateBack"
                    className={ApbBigiconStyleClass.ChevronLeftMed}
                    onClick={() => {
                      ApbAutoSave
                        ? alertDialogforBack()
                        : navType == "ATP"
                          ? props.handleclick("ActivityPlan")
                          : props.handleclick(
                            "ActivityDeliveryPlan",
                            Apb_ActivityPlanId
                          );
                    }}
                  />
                ) : null}
                <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
                  Production board - Activity planner
                </Label>
              </div>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                // marginBottom: 20,
                color: "#2392b2",
                flexWrap: "wrap",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  marginTop: "20px",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                  }}
                  className="toggleWrapper"
                >
                  <label
                    htmlFor="toggle"
                    className={styles.toggle}
                    onChange={(ev) => {
                      if (!Apb_ActivityPlanId) {
                        if (ApbAutoSave) {
                          confirm(
                            "You have unsaved changes, are you sure you want to leave?"
                          )
                            ? setApbChecked(!ApbChecked)
                            : null;
                        } else {
                          setApbChecked(!ApbChecked);
                        }
                      }
                    }}
                  >
                    {!ApbChecked ? (
                      <input type="checkbox" checked id="toggle" />
                    ) : (
                      <input type="checkbox" id="toggle" />
                    )}
                    <span className={styles.slider}>
                      <p>Annual Plan</p>
                      <p>Activity Planner</p>
                    </span>
                  </label>
                </div>
                {!Apb_ActivityPlanId && ApbWeek == Apb_WeekNumber ? (
                  <div>
                    <PrimaryButton
                      text="Ad hoc task"
                      className={ApbbuttonStyleClass.buttonPrimary}
                      onClick={(_) => {
                        let adhocItem = {
                          RefId: ApbData.length + 1,
                          ID: 0,
                          StartDate: new Date(),
                          EndDate: new Date(),
                          Source: "Ad hoc",
                          Project: "",
                          Product: "",
                          Title: "",
                          PlannedHours: 0,
                          Monday: "0",
                          Tuesday: "0",
                          Wednesday: "0",
                          Thursday: "0",
                          Friday: "0",

                          Saturday: "0",

                          Sunday: "0",

                          ActualHours: 0,
                          DeveloperId: loggeduserid,
                          DeveloperEmail: loggeduseremail,
                          Week: Apb_WeekNumber,
                          Year: Apb_Year,
                          Status: null,
                          Lessons: "",
                          Steps: "",
                          ActivityPlanID: null,
                          ActivityDeliveryPlanID: null,
                          ADPActualHours: null,
                          UnPlannedHours: false,
                          PHWeek: null,
                          Onchange: true,
                        };
                        setApbShowMessage(ApbErrorStatus);
                        setApbAdhocPopup({
                          visible: true,
                          isNew: true,
                          value: adhocItem,
                        });
                      }}
                    />
                  </div>
                ) : null}
                <div className={ApbProjectInfo}>
                  <Label className={ApblabelStyles.titleLabel}>
                    Current week :
                  </Label>
                  <Label
                    className={ApblabelStyles.labelValue}
                    style={{ maxWidth: 500 }}
                  >
                    {Apb_WeekNumber}
                  </Label>
                </div>
                <div className={ApbProjectInfo}>
                  <Label className={ApblabelStyles.titleLabel}>
                    Current year :
                  </Label>
                  <Label
                    className={ApblabelStyles.labelValue}
                    style={{ maxWidth: 500 }}
                  >
                    {Apb_Year}
                  </Label>
                </div>
                <div className={ApbProjectInfo}>
                  <Label className={ApblabelStyles.titleLabel}>
                    Actual hrs/ Planned hrs :
                  </Label>
                  <Label
                    className={ApblabelStyles.labelValue}
                    style={{ maxWidth: 500 }}
                  >
                    {sumOfActualHours()} / {sumOfHours()}
                  </Label>
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  marginTop: "20px",
                }}
              >
                <div
                  className={ApbProjectInfo}
                  style={{
                    marginRight: "20px",
                    marginTop: "-24px",
                    transform: "translateY(12px)",
                  }}
                >
                  <Label className={ApblabelStyles.NORLabel}>
                    Number of records:{" "}
                    <b style={{ color: "#038387" }}>{ApbFilterData.length}</b>
                  </Label>
                </div>
                {ApbData.length > 0 &&
                  ApbFilterOptions.Developer == loggeduserid ? (
                  <div>
                    {ApbUpdate ? (
                      <div>
                        <PrimaryButton
                          iconProps={cancelIcon}
                          text="Cancel"
                          className={ApbbuttonStyleClass.buttonPrimary}
                          onClick={(_) => {
                            cancelApbData();
                          }}
                        />
                        <PrimaryButton
                          iconProps={saveIcon}
                          text="Save"
                          id="apdBtnSave"
                          className={ApbbuttonStyleClass.buttonSecondary}
                          onClick={(_) => {
                            setApbAutoSave(false);
                            saveApbData();
                          }}
                        />
                      </div>
                    ) : (
                      <div>
                        <PrimaryButton
                          iconProps={editIcon}
                          text="Edit"
                          className={ApbbuttonStyleClass.buttonPrimary}
                          onClick={() => {
                            setApbUpdate(true);
                            setApbAutoSave(true);

                            //Sorting
                            sortApbUpdate = true;
                            setapbColumns(_apbColumns);
                            setApbData(sortApbDataArr);
                            setApbFilterData(sortApbFilterArr);
                            Activitypaginate(1, sortApbFilterArr);
                          }}
                        />
                        <PrimaryButton
                          iconProps={saveIcon}
                          text="Save"
                          disabled={true}
                          onClick={(_) => { }}
                        />
                      </div>
                    )}
                  </div>
                ) : null}
                <Label
                  onClick={() => {
                    generateExcel();
                  }}
                  style={{
                    backgroundColor: "#EBEBEB",
                    padding: "7px 15px",
                    cursor: "pointer",
                    fontSize: "12px",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    borderRadius: "3px",
                    color: "#1D6F42",
                    marginLeft: 10,
                  }}
                >
                  <Icon
                    style={{
                      color: "#1D6F42",
                    }}
                    iconName="ExcelDocument"
                    className={ApbiconStyleClass.export}
                  />
                  Export as XLS
                </Label>
                {false ? (
                  <Icon
                    iconName="PasteAsText"
                    className={ApbiconStyleClass.Apblink}
                    onClick={() => { }}
                  />
                ) : null}
              </div>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
              }}
            >
              <div className={styles.ddSection}>
                <div>
                  <Label styles={ApbLabelStyles}>Section</Label>
                  <Dropdown
                    placeholder="Select an option"
                    options={ApbDropDownOptions.Lessons}
                    selectedKey={ApbFilterOptions.Lessons}
                    styles={
                      ApbFilterOptions.Lessons == "All"
                        ? ApbDropdownStyles
                        : ApbActiveDropdownStyles
                    }
                    onChange={(e, option: any) => {
                      onChangeFilter("Lessons", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label styles={ApbLabelStyles}>Steps</Label>
                  <Dropdown
                    selectedKey={ApbFilterOptions.Steps}
                    placeholder="Select an option"
                    options={ApbDropDownOptions.Steps}
                    styles={
                      ApbFilterOptions.Steps == "All"
                        ? ApbDropdownStyles
                        : ApbActiveDropdownStyles
                    }
                    onChange={(e, option: any) => {
                      onChangeFilter("Steps", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label styles={ApbLabelStyles}>Product</Label>
                  <Dropdown
                    selectedKey={
                      Apb_ActivityPlanId &&
                        ApbFilterData.length > 0 &&
                        ApbFilterData[0].Product
                        ? ApbFilterData[0].Product
                        : ApbFilterOptions.Product
                    }
                    placeholder="Select an option"
                    options={ApbDropDownOptions.Product}
                    styles={
                      ApbFilterOptions.Product == "All"
                        ? ApbDropdownStyles
                        : ApbActiveDropdownStyles
                    }
                    onChange={(e, option: any) => {
                      onChangeFilter("Product", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label styles={ApbLabelStyles}>Name of the deliverable</Label>
                  <Dropdown
                    selectedKey={
                      Apb_ActivityPlanId
                        ? FilterProject
                        : ApbFilterOptions.Project
                    }
                    placeholder="Select an option"
                    options={ApbDropDownOptions.Project}
                    dropdownWidth={"auto"}
                    styles={
                      ApbFilterOptions.Project == "All"
                        ? ApbDropdownStyles
                        : ApbActiveDropdownStyles
                    }
                    onChange={(e, option: any) => {
                      onChangeFilter("Project", option["key"]);
                    }}
                  />
                </div>
                <div style={{ width: "86px" }}>
                  <Label styles={ApbLabelStyles}>Show only</Label>
                  <Dropdown
                    selectedKey={ApbFilterOptions.Showonly}
                    placeholder="Select an option"
                    options={ApbDropDownOptions.Showonly}
                    styles={showonlyDropdownActive}
                    // style={{ width: "0px" }}
                    onChange={(e, option: any) => {
                      onChangeFilter("Showonly", option["key"]);
                    }}
                  />
                </div>
                <div>
                  {/* <Label styles={ApbLabelStyles}>Developer</Label> */}
                  <Dropdown
                    selectedKey={
                      ApbFilterOptions.Showonly == "All"
                        ? ApbFilterOptions.Developer
                        : loggeduserid
                    }
                    placeholder="Select an option"
                    options={
                      ApbFilterOptions.Showonly == "Mine"
                        ? ApbDropDownOptions.DeveloperMine
                        : ApbDropDownOptions.Developer
                    }
                    styles={ApbActiveDropdownStyles}
                    style={{ marginTop: 25 }}
                    onChange={(e, option: any) => {
                      onChangeFilter("Developer", option["key"]);
                    }}
                  />
                </div>
                {/* <div>
                  <Label styles={ApbLabelStyles}>Week</Label>
                  <Dropdown
                    selectedKey={ApbFilterOptions.Week}
                    placeholder="Select an option"
                    options={ApbDropDownOptions.Week}
                    styles={ApbActiveDropdownStyles}
                    onChange={(e, option: any) => {
                      onChangeFilter("Week", option["key"]);
                    }}
                  />
                </div> */}
                <div>
                  <Label styles={ApbShortLabelStyles}>Week</Label>
                  <Dropdown
                    selectedKey={ApbFilterOptions.WeekNumber}
                    placeholder="Select an option"
                    options={ApbDropDownOptions.WeekNumber}
                    styles={ApbActiveShortDropdownStyles}
                    onChange={(e, option: any) => {
                      onChangeFilter("WeekNumber", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label styles={ApbShortLabelStyles}>Year</Label>
                  <Dropdown
                    selectedKey={ApbFilterOptions.Year}
                    placeholder="Select an option"
                    options={ApbDropDownOptions.Year}
                    styles={ApbActiveShortDropdownStyles}
                    onChange={(e, option: any) => {
                      onChangeFilter("Year", option["key"]);
                    }}
                  />
                </div>

                <div>
                  <div>
                    <Icon
                      iconName="Refresh"
                      title="Click to reset"
                      className={ApbiconStyleClass.refresh}
                      onClick={() => {
                        if (ApbAutoSave) {
                          if (
                            confirm(
                              "You have unsaved changes, are you sure you want to leave?"
                            )
                          ) {
                            setApbWeek(Apb_WeekNumber);
                            setApbYear(Apb_Year);
                            setApbFilterOptions({ ...ApbFilterKeys });

                            if (Apb_ActivityPlanId) {
                              setApbData([...ApbMasterData]);
                              sortApbDataArr = ApbMasterData;

                              let ApbFilter = ActivityProductionBoardFilter(
                                [...ApbMasterData],
                                ApbFilterKeys
                              );
                              setApbFilterData(ApbFilter);
                              sortApbFilterArr = ApbFilter;
                              Activitypaginate(1, ApbFilter);
                              setApbUpdate(false);
                              sortApbUpdate = false;

                              setapbColumns(_apbColumns);
                              getCurrentApbData(
                                Apb_WeekNumber,
                                Apb_Year,
                                ApbFilterKeys
                              );
                            } else {
                              setApbUpdate(false);
                              sortApbUpdate = false;
                              setapbColumns(_apbColumns);
                              getApbData(
                                loggeduserid,
                                Apb_WeekNumber,
                                Apb_Year,
                                ApbFilterKeys
                              );
                            }
                          }
                        } else {
                          setApbWeek(Apb_WeekNumber);
                          setApbYear(Apb_Year);
                          setApbFilterOptions({ ...ApbFilterKeys });

                          if (Apb_ActivityPlanId) {
                            setApbData([...ApbMasterData]);
                            sortApbDataArr = ApbMasterData;

                            let ApbFilter = ActivityProductionBoardFilter(
                              [...ApbMasterData],
                              ApbFilterKeys
                            );
                            setApbFilterData(ApbFilter);
                            sortApbFilterArr = ApbFilter;
                            Activitypaginate(1, ApbFilter);
                            setApbUpdate(false);

                            setapbColumns(_apbColumns);
                            getCurrentApbData(
                              Apb_WeekNumber,
                              Apb_Year,
                              ApbFilterKeys
                            );
                          } else {
                            setApbUpdate(false);
                            sortApbUpdate = false;
                            setapbColumns(_apbColumns);
                            getApbData(
                              loggeduserid,
                              Apb_WeekNumber,
                              Apb_Year,
                              ApbFilterKeys
                            );
                          }
                        }
                      }}
                    />
                  </div>
                </div>
              </div>
              {/* <div
            className={ApbProjectInfo}
            style={{ marginLeft: "20px", transform: "translateY(12px)" }}
          >
            <Label className={ApblabelStyles.NORLabel}>
              Number of records:{" "}
              <b style={{ color: "#038387" }}>{ApbFilterData.length}</b>
            </Label>
          </div> */}
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "flex-end",
                marginTop: 10,
                fontSize: 13.5,
                color: "#323130",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
              >
                <span
                  style={{
                    backgroundColor: "#038387",
                    width: 14,
                    height: 14,
                    borderRadius: 4,
                    marginRight: 6,
                  }}
                ></span>
                Planned hours
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  marginLeft: 10,
                }}
              >
                <span
                  style={{
                    backgroundColor: "#FAA332",
                    width: 14,
                    height: 14,
                    borderRadius: 4,
                    marginRight: 6,
                  }}
                ></span>
                Unplanned hours
              </div>
            </div>
          </div>
          {!ApbChecked ? (
            <div style={{ marginTop: "10px" }}>
              <DetailsList
                items={ApbDisplayData}
                columns={sortApbUpdate ? _apbColumns : apbColumns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                // styles={gridStyles}
                styles={{ root: { width: "100%" } }}
              />
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  margin: "20px 0",
                }}
              >
                {ApbFilterData.length > 0 ? (
                  <Pagination
                    currentPage={ApbcurrentPage}
                    totalPages={
                      ApbFilterData.length > 0
                        ? Math.ceil(ApbFilterData.length / totalPageItems)
                        : 1
                    }
                    onChange={(page) => {
                      Activitypaginate(page, ApbFilterData);
                    }}
                  />
                ) : (
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "center",
                      marginTop: "15px",
                    }}
                  >
                    <Label style={{ color: "#2392B2" }}>
                      No data Found !!!
                    </Label>
                  </div>
                )}
              </div>
            </div>
          ) : (
            props.handleclick(
              "ProductionBoard",
              pbSwitchID,
              pbSwitchType,
              Apb_ActivityPlanId ? Apb_ActivityPlanId + "-" + navType : null
            )
          )}

          <Modal isOpen={ApbModalBoxVisibility} isBlocking={false}>
            <div style={{ padding: "30px 20px" }}>
              <div
                style={{
                  fontSize: 24,
                  textAlign: "center",
                  color: "#2392B2",
                  fontWeight: "600",
                  marginBottom: "20px",
                }}
              >
                Document review
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "flex-start",
                }}
              >
                <div>
                  <Dropdown
                    required={true}
                    errorMessage={AdrPBShowMessage.Request}
                    label="Request"
                    placeholder="Select an option"
                    options={ApbModalBoxDropDownOptions.Request}
                    styles={ApbModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      AdrPBAddOnchange("Request", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label
                    required={true}
                    style={{
                      transform: "translate(20px, 10px)",
                    }}
                  >
                    Request to
                  </Label>
                  <NormalPeoplePicker
                    className={ApbModalBoxPP}
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={1}
                    styles={{
                      root: {
                        width: 300,
                        margin: "10px 20px",
                        selectors: {
                          ".ms-BasePicker-text": {
                            height: 36,
                            padding: "3px 10px",
                            border: "1px solid black",
                            borderRadius: 4,
                          },
                        },
                        ".ms-Persona-primaryText": { fontWeight: 600 },
                      },
                    }}
                    onChange={(selectedUser) => {
                      selectedUser.length != 0
                        ? AdrPBAddOnchange("Requestto", selectedUser[0]["ID"])
                        : AdrPBAddOnchange("Requestto", "");
                    }}
                  />
                  <Label
                    style={{
                      transform: "translate(20px, 10px)",
                      color: "#a4262c",
                      fontSize: 12,
                      fontWeight: 400,
                      paddingTop: 5,
                      marginTop: -20,
                    }}
                  >
                    {AdrPBShowMessage.Requestto}
                  </Label>
                </div>
                <div>
                  <Label
                    style={{
                      transform: "translate(20px, 10px)",
                    }}
                  >
                    Email (cc)
                  </Label>
                  <NormalPeoplePicker
                    className={ApbModalBoxPP}
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={5}
                    styles={{
                      root: {
                        width: 300,
                        margin: "10px 20px",
                        selectors: {
                          ".ms-BasePicker-text": {
                            padding: "3px 10px",
                            border: "1px solid black",
                            borderRadius: 4,
                            maxHeight: "60px",
                            overflowX: "hidden",
                            "::after": {
                              border: "none",
                            },
                          },
                        },
                        ".ms-Persona-primaryText": {
                          fontWeight: 600,
                          border: "none",
                        },
                      },
                    }}
                    onChange={(selectedUser) => {
                      let selectedId = selectedUser.map((su) => su["ID"]);
                      selectedUser.length != 0
                        ? AdrPBAddOnchange("Emailcc", selectedId)
                        : AdrPBAddOnchange("Emailcc", "");
                    }}
                  />
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "flex-start",
                }}
              >
                <div>
                  <TextField
                    label="Name of the deliverable"
                    placeholder="Add name of the deliverable"
                    defaultValue={ApbDocumentReview.Project}
                    disabled={true}
                    styles={ApbTxtBoxStyles}
                    className={styles.projectField}
                    onChange={(e, value: string) => { }}
                  />
                </div>
                <div>
                  <Dropdown
                    label="Document type"
                    required={true}
                    errorMessage={AdrPBShowMessage.Documenttype}
                    placeholder="Select an option"
                    options={ApbModalBoxDropDownOptions.Documenttype}
                    styles={ApbModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      AdrPBAddOnchange("Documenttype", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <TextField
                    label="Link"
                    placeholder="Add link"
                    errorMessage={AdrPBShowMessage.Link}
                    required={true}
                    styles={ApbTxtBoxStyles}
                    onChange={(e, value: string) => {
                      AdrPBAddOnchange("Link", value);
                    }}
                  />
                </div>
              </div>

              <div
              >
                <div>
                  {
                    documentLinkStatus == "incorrect" ? <p style={{ color: "red", textAlign: "center" }}> Incorrect or external link.

                      <a
                        href="https://ggsaus.sharepoint.com/SiteAssets/2023-08-03_15-33-36.png?csf=1&web=1&e=74djHF&cid=ecee1f5a-6654-4a66-89f1-d78b67502786"
                        target="_blank"
                        rel="noopener noreferrer"
                      >
                        Learn how to copy a link
                      </a>
                    </p> :
                      <p style={{ display: "none" }}> </p>
                  }

                </div>

              </div>

              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "flex-start",
                }}
              >
                <div>
                  <TextField
                    label="Comments"
                    placeholder="Add Comments"
                    multiline
                    rows={5}
                    resizable={false}
                    styles={ApbMultiTxtBoxStyles}
                    onChange={(e, value: string) => {
                      AdrPBAddOnchange("Comments", value);
                    }}
                  />
                </div>
                <div
                  style={{
                    marginTop: 30,
                    marginLeft: 20,
                    position: "relative",
                  }}
                >
                  <Toggle
                    label={
                      <div
                        style={{
                          position: "absolute",
                          left: "0",
                          top: "0",
                          width: "200px",
                        }}
                      >
                        Confidential
                      </div>
                    }
                    inlineLabel
                    style={{ transform: "translateX(100px)" }}
                    onChange={(ev) => {
                      AdrPBAddOnchange(
                        "Confidential",
                        !ApbDocumentReview.Confidential
                      );
                    }}
                  />
                </div>
                <div
                  style={{
                    marginTop: 30,
                    marginLeft: 99,
                    position: "relative",
                  }}
                >
                  <Toggle
                    label={
                      <div
                        style={{
                          position: "absolute",
                          left: "0",
                          top: "0",
                          width: "200px",
                        }}
                      >
                        External Link
                      </div>
                    }
                    inlineLabel
                    style={{ transform: "translateX(100px)" }}
                    onChange={(ev) => {
                      AdrPBAddOnchange(
                        "IsExternalAllow",
                        !ApbDocumentReview.IsExternalAllow
                      );
                    }}
                  />
                </div>
              </div>
              <div className={styles.apModalBoxButtonSection}>
                <button
                  className={styles.apModalBoxSubmitBtn}
                  onClick={(_) => {
                    drPBValidationFunction();
                  }}
                  style={{ display: "flex" }}
                >
                  {ApbButtonLoader ? (
                    <Spinner />
                  ) : (
                    <span>
                      <Icon
                        iconName="Save"
                        style={{ position: "relative", top: 3, left: -8 }}
                      />
                      {"Submit"}
                    </span>
                  )}
                </button>
                <button
                  className={styles.apModalBoxBackBtn}
                  onClick={(_) => {
                    setApbModalBoxVisibility(false);
                  }}
                >
                  <span>
                    <Icon
                      iconName="Cancel"
                      style={{ position: "relative", top: 3, left: -8 }}
                    />
                    Close
                  </span>
                </button>
              </div>
            </div>
          </Modal>

          {/* AdhocTask */}
          <Modal isOpen={ApbAdhocPopup.visible} isBlocking={false}>
            <div style={{ padding: "30px 20px" }}>
              <div
                style={{
                  fontSize: 24,
                  textAlign: "center",
                  color: "#2392B2",
                  fontWeight: "600",
                  marginBottom: "20px",
                }}
              >
                Ad hoc task
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "flex-start",
                }}
              >
                <div>
                  <Dropdown
                    label="Type"
                    placeholder="Select a type"
                    required={true}
                    options={ApbModalBoxDropDownOptions.Type}
                    errorMessage={ApbShowMessage.Type}
                    selectedKey={ApbAdhocPopup.value["Title"]}
                    styles={ApbModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      ApbAddOnchange("Title", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <DatePicker
                    label="Start date"
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    styles={ApbModalBoxDatePickerStyles}
                    formatDate={dateFormater}
                    value={ApbAdhocPopup.value["StartDate"]}
                    onSelectDate={(value: any) => {
                      ApbAddOnchange("StartDate", value);
                    }}
                  />
                </div>
                <div>
                  <DatePicker
                    label="End date"
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    styles={ApbModalBoxDatePickerStyles}
                    formatDate={dateFormater}
                    value={ApbAdhocPopup.value["EndDate"]}
                    onSelectDate={(value: any) => {
                      ApbAddOnchange("EndDate", value);
                    }}
                  />
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  // justifyContent: "flex-start",
                  justifyContent: "space-between",
                }}
              >
                <div>
                  <Dropdown
                    label="Name of the deliverable"
                    required={true}
                    placeholder="Select name of the deliverable"
                    style={{ width: "642px" }}
                    options={ApbModalBoxDropDownOptions.Project}
                    errorMessage={ApbShowMessage.Project}
                    selectedKey={ApbAdhocPopup.value["Project"]}
                    styles={ApbModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      ApbAddOnchange("Project", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Dropdown
                    label="Product"
                    required={true}
                    placeholder="Select a product"
                    options={ApbModalBoxDropDownOptions.Product}
                    errorMessage={ApbShowMessage.Product}
                    selectedKey={ApbAdhocPopup.value["Product"]}
                    styles={ApbModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      console.log(option);
                      ApbAddOnchange("Product", option["text"]);
                    }}
                  />
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "flex-start",
                }}
              >
                <div>
                  <TextField
                    label="Section"
                    placeholder="Add section"
                    errorMessage={ApbShowMessage.Lessons}
                    value={ApbAdhocPopup.value["Lessons"]}
                    required={true}
                    styles={ApbTxtBoxStyles}
                    className={styles.projectField}
                    onChange={(e, value: string) => {
                      ApbAddOnchange("Lessons", value);
                    }}
                  />
                </div>
                <div>
                  <TextField
                    label="Task"
                    placeholder="Add task"
                    errorMessage={ApbShowMessage.Steps}
                    value={ApbAdhocPopup.value["Steps"]}
                    required={true}
                    styles={ApbTxtBoxStyles}
                    className={styles.projectField}
                    onChange={(e, value: string) => {
                      ApbAddOnchange("Steps", value);
                    }}
                  />
                </div>
                <div>
                  <TextField
                    label="Hours"
                    placeholder="Add hours"
                    errorMessage={ApbShowMessage.PlannedHours}
                    value={ApbAdhocPopup.value["PlannedHours"]}
                    required={true}
                    styles={ApbTxtBoxStyles}
                    className={styles.projectField}
                    onChange={(e, value: string) => {
                      parseFloat(value)
                        ? ApbAddOnchange("PlannedHours", value)
                        : ApbAddOnchange("PlannedHours", null);
                    }}
                  />
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "flex-start",
                }}
              >
                <div
                  style={{
                    marginTop: 30,
                    marginLeft: 20,
                    position: "relative",
                  }}
                >
                  <Toggle
                    label={
                      <div
                        style={{
                          position: "absolute",
                          left: "0",
                          top: "0",
                          width: "300px",
                        }}
                      >
                        Unplanned hours
                      </div>
                    }
                    inlineLabel
                    checked={ApbAdhocPopup.value["UnPlannedHours"]}
                    style={{ transform: "translateX(100px)", marginLeft: 25 }}
                    onChange={(ev) => {
                      ApbAddOnchange(
                        "UnPlannedHours",
                        !ApbAdhocPopup.value["UnPlannedHours"]
                      );
                    }}
                  />
                </div>
              </div>

              <div className={styles.apModalBoxButtonSection}>
                <button
                  className={styles.apModalBoxSubmitBtn}
                  onClick={(_) => {
                    ApbValidationFunction();
                  }}
                  style={{ display: "flex" }}
                >
                  {ApbButtonLoader ? (
                    <Spinner />
                  ) : (
                    <span>
                      <Icon
                        iconName="Save"
                        style={{ position: "relative", top: 3, left: -8 }}
                      />
                      {ApbAdhocPopup.isNew ? "Submit" : "Update"}
                    </span>
                  )}
                </button>
                <button
                  className={styles.apModalBoxBackBtn}
                  onClick={(_) => {
                    setApbAdhocPopup({
                      visible: false,
                      isNew: true,
                      value: {},
                    });
                  }}
                >
                  <span>
                    <Icon
                      iconName="Cancel"
                      style={{ position: "relative", top: 3, left: -8 }}
                    />
                    Close
                  </span>
                </button>
              </div>
            </div>
          </Modal>

          <div>
            <Modal isOpen={ApbDeletePopup.condition} isBlocking={true}>
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  marginTop: "30px",
                  width: "450px",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "flex-start",
                    flexDirection: "column",
                    marginBottom: "10px",
                  }}
                >
                  <Label className={styles.deletePopupTitle}>
                    Delete deliverable
                  </Label>
                  <Label
                    style={{
                      padding: "5px 20px",
                    }}
                    className={styles.deletePopupDesc}
                  >
                    Are you surewant to delete?
                  </Label>
                </div>
              </div>
              <div className={styles.apDeletePopupBtnSection}>
                <button
                  onClick={(_) => {
                    setApbButtonLoader(true);
                    ApbDeleteItem(ApbDeletePopup.targetId);
                  }}
                  className={styles.apDeletePopupYesBtn}
                >
                  {ApbButtonLoader ? <Spinner /> : "Yes"}
                </button>
                <button
                  onClick={(_) => {
                    setApbDeletePopup({ condition: false, targetId: 0 });
                  }}
                  className={styles.apDeletePopupNoBtn}
                >
                  No
                </button>
              </div>
            </Modal>
          </div>
        </div>
      )}
    </>
  );
}

export default ActivityProductionBoard;
