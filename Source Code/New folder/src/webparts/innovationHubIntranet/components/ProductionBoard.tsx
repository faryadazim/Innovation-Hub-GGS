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
import { IDatePickerStyles, IDetailsListStyles } from "office-ui-fabric-react";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { indexOf } from "lodash";

//Sorting
let sortPbDataArr = [];
let sortPbFilterArr = [];
let sortPbUpdate = false;
let ProjectDetails = [];
let DateListFormat = "DD/MM/YYYY";
let DatePickerFormat = "YYYY-MM-DDT14:00:00Z";

let gblAPDetails = [];

function ProductionBoard(props: any) {
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;
  let pbSwitchID = props.pbSwitch ? props.pbSwitch.split("-")[0] : null;
  let pbSwitchType = props.pbSwitch ? props.pbSwitch.split("-")[1] : null;

  let Ap_AnnualPlanId = props.AnnualPlanId;
  let navType = props.pageType;

  let Pb_Year = moment().year();
  // let Pb_NextWeekYear = moment().add(1, "week").year();
  // let Pb_LastWeekYear = moment().subtract(1, "week").year();

  let Pb_WeekNumber = moment().isoWeek();
  // let Pb_NextWeekNumber = moment().add(1, "week").isoWeek();
  // let Pb_LastWeekNumber = moment().subtract(1, "week").isoWeek();

  // let thisWeekMonday = moment().day(1).format("YYYY-MM-DD");
  // let thisWeekTuesday = moment().day(2).format("YYYY-MM-DD");
  // let thisWeekWednesday = moment().day(3).format("YYYY-MM-DD");
  // let thisWeekThursday = moment().day(4).format("YYYY-MM-DD");
  // let thisWeekFriday = moment().day(5).format("YYYY-MM-DD");

  let loggeduseremail = props.spcontext.pageContext.user.email;
  // let loggeduseremail = "cromulo@goodtogreatschools.org.au";
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

  const pbFilterKeys = {
    BA: "All",
    Source: "All",
    Product: "All",
    Project: "All",
    Showonly: "Mine",
    WeekNumber: Pb_WeekNumber,
    Year: Pb_Year,
    Week: "This Week",
    Developer: loggeduserid,
  };
  let drPBErrorStatus = {
    Request: "",
    Requestto: "",
    Documenttype: "",
    Link: "",
  };
  let pbErrorStatus = {
    BA: "",
    StartDate: "",
    EndDate: "",
    Project: "",
    Product: "",
    Title: "",
    PlannedHours: "",
  };

  const pbDrpDwnOptns = {
    BA: [{ key: "All", text: "All" }],
    Source: [{ key: "All", text: "All" }],
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
    WeekNumber: [{ key: Pb_WeekNumber, text: Pb_WeekNumber.toString() }],
    Year: [{ key: Pb_Year, text: Pb_Year.toString() }],
    DeveloperMine: [{ key: loggeduserid, text: loggerusername }],
    Developer: [{ key: loggeduserid, text: loggerusername }],
  };
  const pbModalBoxDrpDwnOptns = {
    Request: [],
    Documenttype: [],
    BA: [],
    Project: [],
    Product: [],
  };
  const BAacronymsCollection = [
    {
      Name: "PD Curriculum",
      ShortName: "PDC",
    },
    {
      Name: "PD Professional Learning",
      ShortName: "PDPL",
    },
    {
      Name: "PD School Improvements",
      ShortName: "PDSI",
    },
    {
      Name: "SS Business",
      ShortName: "SSB",
    },
    {
      Name: "SS Publishing",
      ShortName: "SSP",
    },
    {
      Name: "SS Content Creation",
      ShortName: "SSCC",
    },
    {
      Name: "SS Marketing",
      ShortName: "SSM",
    },
    {
      Name: "SS Technology",
      ShortName: "SST",
    },
    {
      Name: "SS Research and Evaluation",
      ShortName: "SSRE",
    },
    {
      Name: "SD School Partnerships",
      ShortName: "SSPSP",
    },
  ];

  //Detail list Columns
  const _pbColumns = [
    {
      key: "Column1",
      name: "BA",
      fieldName: "BA",
      minWidth: 40,
      maxWidth: 40,
      onColumnClick: (ev, column) => {
        !sortPbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) =>
        BAacronymsCollection.filter((BAacronym) => {
          return BAacronym.Name == item.BA;
        })[0].ShortName,
    },
    {
      key: "Column2",
      name: "Start Date",
      fieldName: "StartDate",
      minWidth: 80,
      maxWidth: 90,
      onColumnClick: (ev, column) => {
        !sortPbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => item.StartDate,
    },
    {
      key: "Column3",
      name: "End Date",
      fieldName: "EndDate",
      minWidth: 85,
      maxWidth: 90,
      onColumnClick: (ev, column) => {
        !sortPbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => item.EndDate,
    },
    {
      key: "Column4",
      name: "Source",
      fieldName: "Source",
      minWidth: 60,
      maxWidth: 70,
      onColumnClick: (ev, column) => {
        !sortPbUpdate ? _onColumnClick(ev, column) : null;
      },
    },
    {
      key: "Column5",
      name: "Name of the deliverable",
      fieldName: "Project",
      minWidth: 140,
      maxWidth: 400,
      onColumnClick: (ev, column) => {
        !sortPbUpdate ? _onColumnClick(ev, column) : null;
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
      minWidth: 180,
      maxWidth: 350,
      onColumnClick: (ev, column) => {
        !sortPbUpdate ? _onColumnClick(ev, column) : null;
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
      name: "Activity",
      fieldName: "Title",
      minWidth: 160,
      maxWidth: 400,
      onColumnClick: (ev, column) => {
        !sortPbUpdate ? _onColumnClick(ev, column) : null;
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
      key: "Column8",
      name: "PH/UH",
      fieldName: "PlannedHours",
      minWidth: 65,
      maxWidth: 65,
      onColumnClick: (ev, column) => {
        !sortPbUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => (
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
      ),
    },
    {
      key: "Column9",
      name: "Mon",
      fieldName: "Monday",
      minWidth: 40,
      maxWidth: 40,
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
                    width: 40,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              pbUpdate &&
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
              //   ? pbOnchangeItems(item.RefId, "Monday", e.target.value)
              //   : pbOnchangeItems(item.RefId, "Monday", null);
              pbOnchangeItems(item.RefId, "Monday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column10",
      name: "Tue",
      fieldName: "Tuesday",
      minWidth: 40,
      maxWidth: 40,
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
                    width: 40,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              pbUpdate &&
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
              //   ? pbOnchangeItems(item.RefId, "Tuesday", e.target.value)
              //   : pbOnchangeItems(item.RefId, "Tuesday", null);
              pbOnchangeItems(item.RefId, "Tuesday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column11",
      name: "Wed",
      fieldName: "Wednesday",
      minWidth: 40,
      maxWidth: 40,
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
                    width: 40,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              pbUpdate &&
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
              //   ? pbOnchangeItems(item.RefId, "Wednesday", e.target.value)
              //   : pbOnchangeItems(item.RefId, "Wednesday", null);
              pbOnchangeItems(item.RefId, "Wednesday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column12",
      name: "Thu",
      fieldName: "Thursday",
      minWidth: 40,
      maxWidth: 40,
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
                    width: 40,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              pbUpdate &&
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
              //   ? pbOnchangeItems(item.RefId, "Thursday", e.target.value)
              //   : pbOnchangeItems(item.RefId, "Thursday", null);
              pbOnchangeItems(item.RefId, "Thursday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column13",
      name: "Fri",
      fieldName: "Friday",
      minWidth: 40,
      maxWidth: 40,
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
                    width: 40,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              pbUpdate &&
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
              //   ? pbOnchangeItems(item.RefId, "Friday", e.target.value)
              //   : pbOnchangeItems(item.RefId, "Friday", null);
              pbOnchangeItems(item.RefId, "Friday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column14",
      name: "Sat",
      fieldName: "Saturday",
      minWidth: 40,
      maxWidth: 40,
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
                    width: 40,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              pbUpdate &&
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
              // parseFloat(e.target.value)
              //   ? pbOnchangeItems(item.RefId, "Friday", e.target.value)
              //   : pbOnchangeItems(item.RefId, "Friday", null);
              pbOnchangeItems(item.RefId, "Saturday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column15",
      name: "Sun",
      fieldName: "Sunday",
      minWidth: 40,
      maxWidth: 40,
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
                    width: 40,
                    input: {
                      borderRadius: 4,
                    },
                  },
                },
              },
            }}
            data-id={item.ID}
            disabled={
              pbUpdate &&
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
              // parseFloat(e.target.value)
              //   ? pbOnchangeItems(item.RefId, "Friday", e.target.value)
              //   : pbOnchangeItems(item.RefId, "Friday", null);
              pbOnchangeItems(item.RefId, "Sunday", e.target.value);
            }}
          />
        );
      },
    },
    {
      key: "Column16",
      name: "AH",
      fieldName: "ActualHours",
      minWidth: 50,
      maxWidth: 50,
      onColumnClick: (ev, column) => {
        !sortPbUpdate ? _onColumnClick(ev, column) : null;
      },
    },
    {
      key: "Column17",
      name: "Action",
      fieldName: "",
      minWidth: 80,
      maxWidth: 100,
      onRender: (item, Index) =>
        // item.Week == Pb_WeekNumber &&
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
                drAllitems.AnnualPlanID = item.AnnualPlanID;
                drAllitems.DeliveryPlanID = item.DeliveryPlanID;
                drAllitems.ProductionBoardID = item.ID;
                setpbButtonLoader(false);
                setdrPBShowMessage(drPBErrorStatus);
                setpbDocumentReview(drAllitems);
                setpbModalBoxVisibility(true);
              }}
            />
            {item.Source == "Ad hoc" ? (
              <>
                <Icon
                  iconName="Edit"
                  title="Edit deliverable"
                  className={pbiconStyleClass.edit}
                  onClick={() => {
                    setpbButtonLoader(false);
                    let adhocItem = {
                      RefId: item.RefId,
                      ID: item.ID,
                      BA: item.BA,
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
                      AnnualPlanID: item.AnnualPlanID,
                      ProductId: item.ProductId,
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
                      NotApplicable: item.NotApplicable,
                      NotApplicableManager: item.NotApplicableManager,
                      Week: item.Week,
                      Year: item.Year,
                      DeliveryPlanID: item.DeliveryPlanID,
                      DPActualHours: item.DPActualHours,
                      UnPlannedHours: item.UnPlannedHours,
                      Status: item.Status,
                      Onchange: item.Onchange,
                    };
                    setpbShowMessage(pbErrorStatus);
                    setpbAdhocPopup({
                      visible: true,
                      isNew: false,
                      value: adhocItem,
                    });
                  }}
                />
                <Icon
                  iconName="Delete"
                  title="Delete deliverable"
                  className={pbiconStyleClass.delete}
                  onClick={() => {
                    setpbButtonLoader(false),
                      setpbDeletePopup({ condition: true, targetId: item.ID });
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
              cursor: "default",
            }}
            onClick={(_) => { }}
          />
        ) :
          (
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
  const pbModalBoxDatePickerStyles: Partial<IDatePickerStyles> = {
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
      // overflowX: "hidden",
    },
  };
  const pbLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
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
  const pbProjectInfo = mergeStyles({
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    margin: "0 10px",
  });
  const pblabelStyles = mergeStyleSets({
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
  const pbBigiconStyleClass = mergeStyleSets({
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
  const pbbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const pbbuttonStyleClass = mergeStyleSets({
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
      pbbuttonStyle,
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
      pbbuttonStyle,
    ],
  });
  const pbiconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const pbiconStyleClass = mergeStyleSets({
    link: [{ color: "blue", margin: "0 0" }, pbiconStyle],
    delete: [{ color: "#CB1E06", margin: "0 0px" }, pbiconStyle],
    edit: [{ color: "#2392B2", margin: "0 7px" }, pbiconStyle],
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
    pblink: [
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
  const pbActiveShortDropdownStyles: Partial<IDropdownStyles> = {
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
      maxHeight: "300px !important",
    },
  };
  const pbShortLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 75,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const pbDropdownStyles: Partial<IDropdownStyles> = {
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
  const pbActiveDropdownStyles: Partial<IDropdownStyles> = {
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

  const pbModalBoxDropdownStyles: Partial<IDropdownStyles> = {
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
  const pbModalBoxDrpDwnCalloutStyles: Partial<IDropdownStyles> = {
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
  const pbTxtBoxStyles: Partial<ITextFieldStyles> = {
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
  const pbMultiTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "640px",
      margin: "10px 20px",
      borderRadius: "4px",
    },
    field: { fontSize: 12, color: "#000" },
  };
  const pbModalBoxPP = mergeStyles({
    width: "300px",
    margin: "10px 20px",
  });

  // useState
  const [pbReRender, setpbReRender] = useState(false);
  const [pbChecked, setpbChecked] = useState(true);
  const [pbUpdate, setpbUpdate] = useState(false);
  const [pbDisplayData, setpbDisplayData] = useState([]);
  const [pbFilterData, setpbFilterData] = useState([]);
  const [pbData, setpbData] = useState([]);
  const [pbMasterData, setpbMasterData] = useState([]);
  const [pbDropDownOptions, setpbDropDownOptions] = useState(pbDrpDwnOptns);
  const [pbFilterOptions, setpbFilterOptions] = useState(pbFilterKeys);
  const [pbcurrentPage, setpbCurrentPage] = useState(currentpage);
  const [pbLoader, setpbLoader] = useState(true);
  const [pbModalBoxVisibility, setpbModalBoxVisibility] = useState(false);
  const [pbButtonLoader, setpbButtonLoader] = useState(false);
  const [pbModalBoxDropDownOptions, setpbModalBoxDropDownOptions] = useState(
    pbModalBoxDrpDwnOptns
  );
  const [pbDocumentReview, setpbDocumentReview] = useState(drAllitems);
  const [drPBShowMessage, setdrPBShowMessage] = useState(drPBErrorStatus);
  const [pbShowMessage, setpbShowMessage] = useState(pbErrorStatus);
  const [pbWeek, setpbWeek] = useState(Pb_WeekNumber);
  const [pbYear, setpbYear] = useState(Pb_Year);
  // const [pbLastweek, setpbLastweek] = useState([]);
  // const [pbNextweek, setpbNextweek] = useState([]);
  const [pbAutoSave, setpbAutoSave] = useState(false);
  const [pbColumns, setpbColumns] = useState(_pbColumns);
  const [apCurrentData, setapCurrentData] = useState([]);
  const [pbAdhocPopup, setpbAdhocPopup] = useState({
    visible: false,
    isNew: true,
    value: {},
  });
  const [documentLinkStatus, setDocumentLinkStatus] = useState("no-checked")
  const [pbDeletePopup, setpbDeletePopup] = useState({
    condition: false,
    targetId: 0,
  });

  // useEffect
  useEffect(() => {
    //getCamelquery();
    getModalBoxOptions();
    getAPDetails();
    // Ap_AnnualPlanId
    //   ? (getCurrentAPData(),
    //     getCurrentPbData(Pb_WeekNumber, Pb_Year, pbFilterKeys))
    //   : getPbData(loggeduserid, Pb_WeekNumber, Pb_Year, pbFilterKeys);
  }, [pbReRender]);

  useEffect(() => {
    if (pbAutoSave && pbUpdate) {
      setTimeout(() => {
        document.getElementById("pbBtnSave").click();
      }, 300000);
    }
  }, [pbAutoSave]);

  window.onbeforeunload = function (e) {
    debugger;
    if (pbAutoSave) {
      let dialogText =
        "You have unsaved changes, are you sure you want to leave?";
      e.returnValue = dialogText;
      return dialogText;
    }
  };

  const alertDialogforBack = () => {
    if (confirm("You have unsaved changes, are you sure you want to leave?")) {
      navType == "AP"
        ? props.handleclick("AnnualPlan")
        : props.handleclick("DeliveryPlan", Ap_AnnualPlanId);
    }
  };

  const getCamelquery = () => {
    let Filtercondition = `
    <View Scope='RecursiveAll'>
<Query>
   <Where>
      <And>
         <Eq>
            <FieldRef Name='Week' />
            <Value Type='Number'>${Pb_WeekNumber}</Value>
         </Eq>
         <And>
            <Eq>
               <FieldRef Name='Year' />
               <Value Type='Number'>${Pb_Year}</Value>
            </Eq>
            <And>
               <Neq>
                  <FieldRef Name='NotApplicable' />
                  <Value Type='Boolean'>true</Value>
               </Neq>
               <And>
                  <Neq>
                     <FieldRef Name='NotApplicableManager' />
                     <Value Type='Boolean'>true</Value>
                  </Neq>
                  <Eq>
                     <FieldRef Name='Developer' />
                     <Value Type='Text'>${loggerusername}</Value>
                  </Eq>
               </And>
            </And>
         </And>
      </And>
   </Where>
</Query>
<ViewFields>
            <FieldRef Name='ID' />
            <FieldRef Name='Title' />
              <FieldRef Name='BA' />
              <FieldRef Name='StartDate' />
              <FieldRef Name='EndDate' />
              <FieldRef Name='Source' />
              <FieldRef Name='AnnualPlanID' />
              <FieldRef Name='Product' />
              <FieldRef Name='PlannedHours' />
              <FieldRef Name='Monday' />
              <FieldRef Name='Tuesday' />
              <FieldRef Name='Wednesday' />
              <FieldRef Name='Thursday' />
              <FieldRef Name='Friday' />
              <FieldRef Name='Saturday' />
              <FieldRef Name='Sunday' />
              <FieldRef Name='ActualHours' />
              <FieldRef Name='Developer' />
              <FieldRef Name='NotApplicable' />
              <FieldRef Name='NotApplicableManager' />
              <FieldRef Name='Week' />
              <FieldRef Name='Year' />
              <FieldRef Name='DeliveryPlanID' />
              <FieldRef Name='DPActualHours' />
              <FieldRef Name='Status' />
            </ViewFields>
            <RowLimit Paged='TRUE'>5000</RowLimit>
          </View>
    `;

    sharepointWeb.lists
      .getByTitle("ProductionBoard")
      .renderListDataAsStream({
        ViewXml: Filtercondition,
      })
      .then((data) => {
        console.log(data.Row, "caml");
      });
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
              pbModalBoxDrpDwnOptns.Request.findIndex((rpb) => {
                return rpb.key == choice;
              }) == -1
            ) {
              pbModalBoxDrpDwnOptns.Request.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch((error) => {
        pbErrorFunction(error, "getModalBoxOptions1");
      });

    //Documenttype Choices
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .fields.getByInternalNameOrTitle("Documenttype")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              pbModalBoxDrpDwnOptns.Documenttype.findIndex((rdt) => {
                return rdt.key == choice;
              }) == -1
            ) {
              pbModalBoxDrpDwnOptns.Documenttype.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch((error) => {
        pbErrorFunction(error, "getModalBoxOptions2");
      });

    //BA Choices
    sharepointWeb.lists
      .getByTitle("ProductionBoard")
      .fields.getByInternalNameOrTitle("BA")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              pbModalBoxDrpDwnOptns.BA.findIndex((rpb) => {
                return rpb.key == choice;
              }) == -1
            ) {
              pbModalBoxDrpDwnOptns.BA.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch((error) => {
        pbErrorFunction(error, "getModalBoxOptions3");
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
              pbModalBoxDropDownOptions.Product.findIndex((productOptn) => {
                return productOptn.text == product.Title;
              }) == -1
            ) {
              if (product.Title != "Not Sure") {
                pbModalBoxDropDownOptions.Product.push({
                  key: product.Id,
                  text: product.Title + " " + product.ProductVersion,
                });
              } else {
                NotSureId = product.Id;
              }
            }
          }
        });
      })
      .then(() => {
        pbModalBoxDropDownOptions.Product.sort(_sortFilterKeys);
        pbModalBoxDropDownOptions.Product.unshift({
          key: NotSureId,
          text: "Not Sure",
        });
      })
      .catch((error) => {
        pbErrorFunction(error, "getModalBoxOptions4");
      });

    //Project Choices AP
    ProjectDetails = [];
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.top(5000)
      .get()
      .then((allPrj) => {
        allPrj.forEach((prj) => {
          if (prj.Title != null) {
            if (
              pbModalBoxDropDownOptions.Project.findIndex((productOptn) => {
                return productOptn.key == prj.Title;
              }) == -1
            ) {
              pbModalBoxDropDownOptions.Project.push({
                key: prj.Title + " " + prj.ProjectVersion,
                text: prj.Title + " " + prj.ProjectVersion,
              });
              ProjectDetails.push({
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
        pbModalBoxDropDownOptions.Project.sort(_sortFilterKeys);
      })
      .catch((error) => {
        pbErrorFunction(error, "getModalBoxOptions5");
      });
    setpbModalBoxDropDownOptions(pbModalBoxDrpDwnOptns);
  };
  const generateExcel = () => {
    let arrExport = pbFilterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Business area", key: "Businessarea", width: 25 },
      { header: "Start date", key: "Startdate", width: 25 },
      { header: "End date", key: "Enddate", width: 25 },
      { header: "Source", key: "Source", width: 25 },
      { header: "Name of the deliverable", key: "POT", width: 25 },
      { header: "Product", key: "Product", width: 60 },
      { header: "Activity", key: "Activity", width: 20 },
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
        Businessarea: item.BA ? item.BA : "",
        Startdate: item.StartDate ? item.StartDate : "",
        Enddate: item.EndDate ? item.EndDate : "",
        Source: item.Source ? item.Source : "",
        POT: item.Project ? item.Project : "",
        Product: item.Product ? item.Product : "",
        Activity: item.Title ? item.Title : "",
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
      "P1",
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
      "P1",
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
          `ProductionBoard-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };

  const getCurrentAPData = () => {
    let _apCurrentData = [];
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.select(
        "*",
        "ProjectOwner/Title",
        "ProjectOwner/Id",
        "ProjectOwner/EMail",
        "ProjectLead/Title",
        "ProjectLead/Id",
        "ProjectLead/EMail",
        "Master_x0020_Project/Title",
        "Master_x0020_Project/Id"
      )
      .expand("ProjectOwner", "ProjectLead", "Master_x0020_Project")
      .filter("ID eq '" + Ap_AnnualPlanId + "' ")
      .top(5000)
      .get()
      .then(async (items) => {
        items.forEach((item) => {
          _apCurrentData.push({
            ID: item.ID,
            Title: item.Title,
            TypeofProject: item.ProjectType,
            ProductId: item.Master_x0020_ProjectId,
            ProductName: item.Master_x0020_Project
              ? item.Master_x0020_Project.Title
              : "",
            DeveloperId: item.ProjectLeadId ? item.ProjectLeadId[0] : null,
            ProjectOwnerId: item.ProjectOwnerId,
            DeveloperEmail: item.ProjectLead ? item.ProjectLead[0].EMail : null,
            ProjectOwnerEmail: item.ProjectOwner
              ? item.ProjectOwner.EMail
              : null,
            DeveloperName: item.ProjectLead ? item.ProjectLead[0].Title : null,
            ProjectOwnerName: item.ProjectOwner
              ? item.ProjectOwner.Title
              : null,
            StartDate: item.StartDate,
            PlannedEndDate: item.PlannedEndDate,
            AllocatedHours: item.AllocatedHours ? item.AllocatedHours : 0,
            BA: item.BusinessArea,
            MultipleDeveloper: item.ProjectLead ? item.ProjectLead : null,
          });
        });

        setapCurrentData([..._apCurrentData]);
      })
      .catch((error) => {
        pbErrorFunction(error, "getCurrentAPData");
      });
  };

  const getAPDetails = () => {
    gblAPDetails = [];
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.select(
        "*",
        "ProjectOwner/Title",
        "ProjectOwner/Id",
        "ProjectOwner/EMail",
        "ProjectLead/Title",
        "ProjectLead/Id",
        "ProjectLead/EMail",
        "Master_x0020_Project/Title",
        "Master_x0020_Project/ProductVersion",
        "Master_x0020_Project/Id"
      )
      .expand("ProjectOwner", "ProjectLead", "Master_x0020_Project")
      .top(5000)
      .get()
      .then(async (items) => {
        items.forEach((item) => {
          gblAPDetails.push({
            ID: item.ID,
            Title: item.Title,
            TypeofProject: item.ProjectType,
            ProductId: item.Master_x0020_ProjectId,
            ProductName: item.Master_x0020_Project
              ? item.Master_x0020_Project.Title
              : "",
            ProductVersion: item.Master_x0020_Project
              ? item.Master_x0020_Project.ProductVersion
              : "",
            DeveloperId: item.ProjectLeadId ? item.ProjectLeadId[0] : null,
            ProjectOwnerId: item.ProjectOwnerId,
            DeveloperEmail: item.ProjectLead ? item.ProjectLead[0].EMail : null,
            ProjectOwnerEmail: item.ProjectOwner
              ? item.ProjectOwner.EMail
              : null,
            DeveloperName: item.ProjectLead ? item.ProjectLead[0].Title : null,
            ProjectOwnerName: item.ProjectOwner
              ? item.ProjectOwner.Title
              : null,
            StartDate: item.StartDate,
            PlannedEndDate: item.PlannedEndDate,
            AllocatedHours: item.AllocatedHours ? item.AllocatedHours : 0,
            BA: item.BusinessArea,
            MultipleDeveloper: item.ProjectLead ? item.ProjectLead : null,
          });
        });
        Ap_AnnualPlanId
          ? getCurrentPbData(Pb_WeekNumber, Pb_Year, pbFilterKeys)
          : getPbData(loggeduserid, Pb_WeekNumber, Pb_Year, pbFilterKeys);
      })
      .catch((error) => {
        pbErrorFunction(error, "getAPDetails");
      });
  };
  const getCurrentPbData = (Week, Year, filterkeys) => {
    setpbLoader(true);
    sharepointWeb.lists
      .getByTitle("ProductionBoard")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,Product/ProductVersion,AnnualPlanID/Title,AnnualPlanID/ProjectVersion,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
      )
      .expand("Developer,Product,AnnualPlanID,FieldValuesAsText")
      .filter(
        "AnnualPlanID eq '" +
        Ap_AnnualPlanId +
        "' and Week eq '" +
        Week +
        "' and Year eq '" +
        Year +
        "' "
      )
      .top(5000)
      .get()
      .then((items) => {
        console.log(items, 'itemssss');

        let _pbAllitems = [];
        let Count = 0;
        items.forEach((item, Index) => {
          let curAPDetails = gblAPDetails.filter((arr) => {
            return arr.ID == item.AnnualPlanIDId;
          });

          _pbAllitems.push({
            RefId: Count++,
            ID: item.ID,
            BA: item.BA,
            StartDate: moment(
              item["FieldValuesAsText"].StartDate,
              DateListFormat
            ).format(DateListFormat),
            EndDate: moment(
              item["FieldValuesAsText"].EndDate,
              DateListFormat
            ).format(DateListFormat),
            Source: item.Source,
            Project: item.AnnualPlanID
              ? item.AnnualPlanID.Title + " " + item.AnnualPlanID.ProjectVersion
              : item.Project +
              " " +
              (item.ProjectVersion ? item.ProjectVersion : "V1"),
            AnnualPlanID: item.AnnualPlanIDId,
            ProductId:
              curAPDetails.length > 0
                ? curAPDetails[0].ProductId
                : item.ProductId,
            Product:
              curAPDetails.length > 0
                ? curAPDetails[0].ProductName +
                " " +
                curAPDetails[0].ProductVersion
                : item.Product
                  ? item.Product.Title +
                  " " +
                  (item.Product.ProductVersion
                    ? item.Product.ProductVersion
                    : "V1")
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
            NotApplicable: item.NotApplicable,
            NotApplicableManager: item.NotApplicableManager,
            Week: item.Week,
            Year: item.Year,
            DeliveryPlanID: item.DeliveryPlanID,
            DPActualHours: item.DPActualHours ? item.DPActualHours : 0,
            UnPlannedHours: item.UnPlannedHours ? item.UnPlannedHours : false,
            Status: item.Status,
            Onchange:
              curAPDetails.length > 0 &&
                curAPDetails[0].ProductId != item.ProductId
                ? true
                : false,
          });
        });

        // if (_pbAllitems.length == 0) {
        getCurrentDpData(Week, Year, _pbAllitems, Count, filterkeys);
        // } else {
        //   let pbOnloadFilter = PBOnloadFilter([..._pbAllitems], filterkeys);
        //   setpbData([...pbOnloadFilter]);
        //   sortPbDataArr = pbOnloadFilter;
        //   setpbMasterData([...pbOnloadFilter]);
        //   let pbFilter = ProductionBoardFilter([...pbOnloadFilter], filterkeys);
        //   reloadFilterOptions([...pbFilter]);
        //   setpbFilterData(pbFilter);
        //   sortPbFilterArr = pbFilter;
        //   paginate(1, pbFilter);
        //   setpbLoader(false);
        // }
      })
      .catch((error) => {
        pbErrorFunction(error, "getCurrentPbData");
      });
  };
  const getCurrentDpData = (Week, Year, data, Count, filterkeys) => {
    sharepointWeb.lists
      .getByTitle("Delivery Plan")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,Product/ProductVersion,AnnualPlanID/Title,AnnualPlanID/ProjectVersion,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
      )
      .expand("Developer,Product,AnnualPlanID,FieldValuesAsText")
      .filter("AnnualPlanID eq '" + Ap_AnnualPlanId + "' ")
      .top(5000)
      .get()
      .then((items) => {
        let _pbAllitems = data;
        let count = Count;
        items.forEach((item, Index) => {
          if (
            _pbAllitems.findIndex((pb) => {
              return pb.DeliveryPlanID == item.ID;
            }) == -1
          ) {
            let curAPDetails = gblAPDetails.filter((arr) => {
              return arr.ID == item.AnnualPlanIDId;
            });

            _pbAllitems.push({
              RefId: count++,
              ID: 0,
              BA: item.BA,
              StartDate: moment(
                item["FieldValuesAsText"].StartDate,
                DateListFormat
              ).format(DateListFormat),
              EndDate: moment(
                item["FieldValuesAsText"].EndDate,
                DateListFormat
              ).format(DateListFormat),
              Source: item.Source,
              Project: item.AnnualPlanID
                ? item.AnnualPlanID.Title +
                " " +
                item.AnnualPlanID.ProjectVersion
                : item.Project +
                " " +
                (item.ProjectVersion ? item.ProjectVersion : "V1"),
              AnnualPlanID: item.AnnualPlanIDId,
              ProductId:
                curAPDetails.length > 0
                  ? curAPDetails[0].ProductId
                  : item.ProductId,
              Product:
                curAPDetails.length > 0
                  ? curAPDetails[0].ProductName +
                  " " +
                  curAPDetails[0].ProductVersion
                  : item.Product
                    ? item.Product.Title +
                    " " +
                    (item.Product.ProductVersion
                      ? item.Product.ProductVersion
                      : "V1")
                    : "",
              Title: item.Title,
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
              NotApplicable: item.NotApplicable,
              NotApplicableManager: item.NotApplicableManager,
              Week: Week,
              Year: Year,
              DeliveryPlanID: item.ID,
              DPActualHours: item.ActualHours ? item.ActualHours : 0,
              UnPlannedHours: item.UnPlannedHours ? item.UnPlannedHours : false,
              Status: null,
              Onchange: false,
            });
          }
        });
        let pbOnloadFilter = PBOnloadFilter([..._pbAllitems], filterkeys);
        setpbData([...pbOnloadFilter]);
        sortPbDataArr = pbOnloadFilter;
        setpbMasterData([...pbOnloadFilter]);
        reloadFilterOptions([...pbOnloadFilter]);
        let pbFilter = ProductionBoardFilter([...pbOnloadFilter], filterkeys);
        setpbFilterData(pbFilter);
        sortPbFilterArr = pbFilter;
        paginate(1, pbFilter);
        setpbLoader(false);
      })
      .catch((error) => {
        pbErrorFunction(error, "getCurrentDpData");
      });
  };
  const getPbData = (devID, Week, Year, filterkeys) => {
    setpbLoader(true);
    sharepointWeb.lists
      .getByTitle("ProductionBoard")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,Product/ProductVersion,AnnualPlanID/Title,AnnualPlanID/ProjectVersion,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
      )
      .expand("Developer,Product,AnnualPlanID,FieldValuesAsText")
      // .filter(
      //   "Week eq '" +
      //     Pb_WeekNumber +
      //     "' and Year eq '" +
      //     Pb_Year +
      //     "' and Developer/EMail eq '" +
      //     loggeduseremail +
      //     "' "
      // )
      // .filter("Week eq '" + Pb_WeekNumber + "' and Year eq '" + Pb_Year + "'")
      .filter(`SPFxFilter eq '${devID}-${Week}-${Year}-false'`)
      .top(5000)
      .get()
      .then((items) => {
        let _pbAllitems = [];
        let Count = 0;
        items.forEach((item, Index) => {
          let curAPDetails = gblAPDetails.filter((arr) => {
            return arr.ID == item.AnnualPlanIDId;
          });

          if (
            //for Deleted AnnualPlan
            (curAPDetails.length > 0 && item.AnnualPlanIDId) ||
            (item.Project && !item.AnnualPlanIDId)
          ) {
            _pbAllitems.push({
              RefId: Count++,
              ID: item.ID,
              BA: item.BA,
              StartDate: moment(
                item["FieldValuesAsText"].StartDate,
                DateListFormat
              ).format(DateListFormat),
              EndDate: moment(
                item["FieldValuesAsText"].EndDate,
                DateListFormat
              ).format(DateListFormat),
              Source: item.Source,
              Project: item.AnnualPlanID
                ? item.AnnualPlanID.Title +
                " " +
                item.AnnualPlanID.ProjectVersion
                : item.Project +
                " " +
                (item.ProjectVersion ? item.ProjectVersion : "V1"),
              AnnualPlanID: item.AnnualPlanIDId,
              ProductId:
                curAPDetails.length > 0
                  ? curAPDetails[0].ProductId
                  : item.ProductId,
              Product:
                curAPDetails.length > 0
                  ? curAPDetails[0].ProductName +
                  " " +
                  curAPDetails[0].ProductVersion
                  : item.Product
                    ? item.Product.Title +
                    " " +
                    (item.Product.ProductVersion
                      ? item.Product.ProductVersion
                      : "V1")
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
              NotApplicable: item.NotApplicable,
              NotApplicableManager: item.NotApplicableManager,
              Week: item.Week,
              Year: item.Year,
              DeliveryPlanID: item.DeliveryPlanID,
              DPActualHours: item.DPActualHours ? item.DPActualHours : 0,
              UnPlannedHours: item.UnPlannedHours ? item.UnPlannedHours : false,
              Status: item.Status,
              Onchange:
                curAPDetails.length > 0 &&
                  curAPDetails[0].ProductId != item.ProductId
                  ? true
                  : false,
            });
          }
        });
        getDpData(Week, Year, _pbAllitems, Count, devID, filterkeys);
      })
      .catch((error) => {
        pbErrorFunction(error, "getPbData");
      });
  };
  const getDpData = (Week, Year, data, Count, devID, filterkeys) => {
    sharepointWeb.lists
      .getByTitle("Delivery Plan")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,Product/ProductVersion,AnnualPlanID/Title,AnnualPlanID/ProjectVersion,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
      )
      .expand("Developer,Product,AnnualPlanID,FieldValuesAsText")
      // .filter("DeveloperId eq '" + loggeduserid + "' ")
      .filter(`SPFxFilter eq '${devID}-0-false'`)
      .top(5000)
      .get()
      .then(async (items) => {
        console.log(items);
        let _pbAllitems = data;
        let count = Count;

        // let _pbAllitems = [];
        // let count = 0;
        // let getAnnualID = records.reduce(function (item, e1) {
        //   var matches = item.filter(function (e2) {
        //     return e1.AnnualPlanIDId === e2.AnnualPlanIDId;
        //   });
        //   if (matches.length == 0) {
        //     item.push(e1);
        //   }
        //   return item;
        // }, []);
        // if (getAnnualID.length > 0) {
        //   await getAnnualID.forEach(async (items) => {
        //     await sharepointWeb.lists
        //       .getByTitle("Delivery Plan")
        //       .items.select(
        //         "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
        //       )
        //       .expand("Developer,Product,AnnualPlanID")
        //       .filter("AnnualPlanID eq '" + items.AnnualPlanIDId + "' ")
        //       .top(5000)
        //       .get()
        //       .then((items) => {
        items.forEach((item) => {
          if (
            _pbAllitems.findIndex((pb) => {
              return pb.DeliveryPlanID == item.ID;
            }) == -1
          ) {
            let curAPDetails = gblAPDetails.filter((arr) => {
              return arr.ID == item.AnnualPlanIDId;
            });
            if (
              //for Deleted AnnualPlan
              (curAPDetails.length > 0 && item.AnnualPlanIDId) ||
              (item.Project && !item.AnnualPlanIDId)
            ) {
              _pbAllitems.push({
                RefId: count++,
                ID: 0,
                BA: item.BA,
                StartDate: moment(
                  item["FieldValuesAsText"].StartDate,
                  DateListFormat
                ).format(DateListFormat),
                EndDate: moment(
                  item["FieldValuesAsText"].EndDate,
                  DateListFormat
                ).format(DateListFormat),
                Source: item.Source,
                Project: item.AnnualPlanID
                  ? item.AnnualPlanID.Title +
                  " " +
                  item.AnnualPlanID.ProjectVersion
                  : item.Project +
                  " " +
                  (item.ProjectVersion ? item.ProjectVersion : "V1"),
                AnnualPlanID: item.AnnualPlanIDId,
                ProductId:
                  curAPDetails.length > 0
                    ? curAPDetails[0].ProductId
                    : item.ProductId,
                Product:
                  curAPDetails.length > 0
                    ? curAPDetails[0].ProductName +
                    " " +
                    curAPDetails[0].ProductVersion
                    : item.Product
                      ? item.Product.Title +
                      " " +
                      (item.Product.ProductVersion
                        ? item.Product.ProductVersion
                        : "V1")
                      : "",
                Title: item.Title,
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
                NotApplicable: item.NotApplicable,
                NotApplicableManager: item.NotApplicableManager,
                Week: Week,
                Year: Year,
                DeliveryPlanID: item.ID,
                DPActualHours: item.ActualHours ? item.ActualHours : 0,
                UnPlannedHours: item.UnPlannedHours
                  ? item.UnPlannedHours
                  : false,
                Status: null,
                Onchange: false,
              });
            }
          }
        });

        let pbOnloadFilter = PBOnloadFilter([..._pbAllitems], filterkeys);
        setpbData([...pbOnloadFilter]);
        sortPbDataArr = pbOnloadFilter;
        setpbMasterData([...pbOnloadFilter]);
        reloadFilterOptions([...pbOnloadFilter]);
        let pbFilter = ProductionBoardFilter([...pbOnloadFilter], filterkeys);
        setpbFilterData(pbFilter);
        sortPbFilterArr = pbFilter;
        paginate(1, pbFilter);
        setpbLoader(false);
        console.log(pbFilter);
      })
      .catch((error) => {
        pbErrorFunction(error, "getDpData");
      });
    // });
    // } else {
    //   setpbLoader(false);
    // }
    // })
    // .catch((error) => {
    //   pbErrorFunction(error, "getAPBList");
    // });
  };
  const savePBData = () => {
    setpbLoader(true);
    let successCount = 0;
    pbData.forEach((pb, Index: number) => {
      let strDWYNA: string = `${pb.DeveloperId}-${pb.Week}-${pb.Year}-${pb.NotApplicable != true && pb.NotApplicableManager != true
        ? false
        : true
        }`;

      let PrjData = ProjectDetails.filter((arr) => {
        return arr.Key == pb.Project;
      });

      let curProjectData = pb.Project ? pb.Project.split("V") : [];
      let curProject =
        curProjectData.length > 0 ? curProjectData[0].trim() : "";

      let PrjVersion = PrjData.length > 0 ? PrjData[0].Version : "V1";
      let PrjTitle = PrjData.length > 0 ? PrjData[0].Title : curProject;

      let requestdata = {
        BA: pb.BA,
        StartDate: pb.StartDate
          ? moment(pb.StartDate, DateListFormat).format("YYYY-MM-DD")
          : null,
        EndDate: pb.EndDate
          ? moment(pb.EndDate, DateListFormat).format("YYYY-MM-DD")
          : null,
        Source: pb.Source ? pb.Source : null,
        AnnualPlanIDId: pb.AnnualPlanID ? pb.AnnualPlanID : null,
        ProductId: pb.ProductId ? pb.ProductId : null,
        Title: pb.Title ? pb.Title : null,
        PlannedHours: pb.PlannedHours ? pb.PlannedHours : null,
        Monday: pb.Monday ? pb.Monday : "0",
        Tuesday: pb.Tuesday ? pb.Tuesday : "0",
        Wednesday: pb.Wednesday ? pb.Wednesday : "0",
        Thursday: pb.Thursday ? pb.Thursday : "0",
        Friday: pb.Friday ? pb.Friday : "0",
        Saturday: pb.Saturday ? pb.Saturday : "0",

        Sunday: pb.Sunday ? pb.Sunday : "0",

        ActualHours: pb.ActualHours ? pb.ActualHours : 0,
        DeveloperId: pb.DeveloperId ? pb.DeveloperId : null,
        Week: pb.Week,
        Year: pb.Year,
        NotApplicable: pb.NotApplicable,
        NotApplicableManager: pb.NotApplicableManager,
        DeliveryPlanID: pb.DeliveryPlanID,
        DPActualHours: pb.DPActualHours ? pb.DPActualHours : 0,
        UnPlannedHours: pb.UnPlannedHours ? pb.UnPlannedHours : false,
        Status: pb.Status ? pb.Status : null,
        AnnualPlanIDNumber: pb.AnnualPlanID,
        Project: PrjTitle,
        ProjectVersion: PrjVersion,
        SPFxFilter: strDWYNA,
      };
      let AH =
        parseFloat(pb.DRActualHours ? pb.DRActualHours : 0) +
        parseFloat(pb.ActualHours ? pb.ActualHours : 0);

      // Adhoc task
      let strDSNA: string = `${pb.DeveloperId}-0-false`;
      let requestdataforDP;
      if (pb.Source == "Ad hoc") {
        requestdataforDP = {
          AnnualPlanIDId: null,
          Project: PrjTitle,
          Source: pb.Source,
          Title: pb.Title ? pb.Title : null,
          StartDate: pb.StartDate
            ? moment(pb.StartDate, DateListFormat).format("YYYY-MM-DD")
            : null,
          EndDate: pb.EndDate
            ? moment(pb.EndDate, DateListFormat).format("YYYY-MM-DD")
            : null,
          DeveloperId: pb.DeveloperId ? pb.DeveloperId : null,
          ManagerId: null,
          NotApplicable: null,
          NotApplicableManager: null,
          Status: "Scheduled",
          PlannedHours: pb.PlannedHours ? pb.PlannedHours : 0,
          ProductId: pb.ProductId ? pb.ProductId : null,
          Week: pb.Week,
          Year: pb.Year,
          OrderId: 0,
          BA: pb.BA ? pb.BA : null,
          AnnualPlanIDNumber: null,
          ActualHours: AH,
          SPFxFilter: strDSNA,
          ProjectVersion: PrjVersion,
          UnPlannedHours: pb.UnPlannedHours ? pb.UnPlannedHours : false,
        };
      } else {
        requestdataforDP = {
          ActualHours: AH,
        };
      }

      if (pb.ID != 0 && pb.Onchange == true) {
        sharepointWeb.lists
          .getByTitle("ProductionBoard")
          .items.getById(pb.ID)
          .update(requestdata)
          .then(() => {
            if (pb.DeliveryPlanID) {
              sharepointWeb.lists
                .getByTitle("Delivery Plan")
                .items.getById(pb.DeliveryPlanID)
                .update(requestdataforDP)
                .then((e) => { })
                .catch((error) => {
                  pbErrorFunction(error, "savePBData1");
                });
            }

            successCount++;
            if (successCount == pbData.length) {
              setpbLoader(false);
              setpbMasterData([...pbData]);
              setpbUpdate(!pbUpdate);
              sortPbUpdate = !pbUpdate;
              AddSuccessPopup();
            }
          })
          .catch((error) => {
            pbErrorFunction(error, "savePBData2");
          });
      } else if (pb.ID == 0 && pb.Onchange == true) {
        if (pb.Source == "Ad hoc" && pb.DeliveryPlanID == null) {
          sharepointWeb.lists
            .getByTitle("Delivery Plan")
            .items.add(requestdataforDP)
            .then((ev) => {
              requestdata.DeliveryPlanID = ev.data.Id;
              pbData[Index].DeliveryPlanID = ev.data.Id;

              sharepointWeb.lists
                .getByTitle("ProductionBoard")
                .items.add(requestdata)
                .then((e) => {
                  if (pb.DeliveryPlanID) {
                    sharepointWeb.lists
                      .getByTitle("Delivery Plan")
                      .items.getById(pb.DeliveryPlanID)
                      .update({ ActualHours: AH })
                      .then((e) => { })
                      .catch((error) => {
                        pbErrorFunction(error, "savePBData3");
                      });
                  }
                  successCount++;
                  pbData[Index].ID = e.data.ID;
                  if (successCount == pbData.length) {
                    setpbLoader(false);
                    setpbData([...pbData]);
                    sortPbDataArr = pbData;
                    setpbMasterData([...pbData]);
                    setpbUpdate(!pbUpdate);
                    sortPbUpdate = !pbUpdate;
                    AddSuccessPopup();
                  }
                })
                .catch((error) => {
                  pbErrorFunction(error, "savePBData4");
                });
            })
            .catch((error) => {
              pbErrorFunction(error, "savePBData4");
            });
        } else {
          sharepointWeb.lists
            .getByTitle("ProductionBoard")
            .items.add(requestdata)
            .then((e) => {
              if (pb.DeliveryPlanID) {
                sharepointWeb.lists
                  .getByTitle("Delivery Plan")
                  .items.getById(pb.DeliveryPlanID)
                  .update({ ActualHours: AH })
                  .then((e) => { })
                  .catch((error) => {
                    pbErrorFunction(error, "savePBData3");
                  });
              }
              successCount++;
              pbData[Index].ID = e.data.ID;
              if (successCount == pbData.length) {
                setpbLoader(false);
                setpbData([...pbData]);
                sortPbDataArr = pbData;
                setpbMasterData([...pbData]);
                setpbUpdate(!pbUpdate);
                sortPbUpdate = !pbUpdate;
                AddSuccessPopup();
              }
            })
            .catch((error) => {
              pbErrorFunction(error, "savePBData4");
            });
        }
      } else {
        successCount++;
        if (successCount == pbData.length) {
          setpbUpdate(!pbUpdate);
          sortPbUpdate = !pbUpdate;
          setpbLoader(false);
          AddSuccessPopup();
        }
      }
    });
  };
  const savePBDRData = () => {
    let requestdata = {
      Link: pbDocumentReview.Link,
      Request: pbDocumentReview.Request ? pbDocumentReview.Request : null,
      RequesttoId: pbDocumentReview.Requestto
        ? pbDocumentReview.Requestto
        : null,
      EmailccId: pbDocumentReview.Emailcc
        ? { results: pbDocumentReview.Emailcc }
        : { results: [] },
      Project: pbDocumentReview.Project ? pbDocumentReview.Project : null,
      Documenttype: pbDocumentReview.Documenttype
        ? pbDocumentReview.Documenttype
        : null,
      Comments: pbDocumentReview.Comments ? pbDocumentReview.Comments : null,
      Confidential: pbDocumentReview.Confidential,
      IsExternalAllow: pbDocumentReview.IsExternalAllow,
      Product: pbDocumentReview.Product ? pbDocumentReview.Product : null,
      AnnualPlanID: pbDocumentReview.AnnualPlanID
        ? pbDocumentReview.AnnualPlanID
        : null,
      DeliveryPlanID: pbDocumentReview.DeliveryPlanID
        ? pbDocumentReview.DeliveryPlanID
        : null,
      ProductionBoardID: pbDocumentReview.ProductionBoardID
        ? pbDocumentReview.ProductionBoardID
        : null,
      DRPageName: "Annual Plan",
    };
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .items.add(requestdata)
      .then((e) => {
        if (pbDocumentReview.ProductionBoardID) {
          sharepointWeb.lists
            .getByTitle("ProductionBoard")
            .items.getById(pbDocumentReview.ProductionBoardID)
            .update({ Status: "Pending" })
            .then(() => {
              let Index = pbData.findIndex(
                (obj) => obj.ID == pbDocumentReview.ProductionBoardID
              );
              pbData[Index].Status = "Pending";
              setpbData([...pbData]);
              sortPbDataArr = pbData;
              setDocumentLinkStatus("no-checked")
            })
            .catch((error) => {
              setDocumentLinkStatus("no-checked")
              pbErrorFunction(error, "savePBDRData1");
            });
        }
        setpbModalBoxVisibility(false);
        AddDRSuccessPopup();
      })
      .catch((error) => {
        pbErrorFunction(error, "savePBDRData2");
      });
  };
  const cancelPBData = () => {
    setDocumentLinkStatus("no-checked")
    // setpbFilterOptions({ ...pbFilterKeys });
    setpbData([...pbMasterData]);
    sortPbDataArr = pbMasterData;
    setpbUpdate(false);
    sortPbUpdate = false;
    let pbFilter = ProductionBoardFilter([...pbMasterData], pbFilterKeys);
    setpbFilterData(pbFilter);
    sortPbFilterArr = pbFilter;
    paginate(1, pbFilter);
    setpbAutoSave(false);
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;
    tempArrReload.forEach((item) => {
      if (
        pbDrpDwnOptns.BA.findIndex((BA) => {
          return BA.key == item.BA;
        }) == -1 &&
        item.BA
      ) {
        pbDrpDwnOptns.BA.push({
          key: item.BA,
          text: item.BA,
        });
      }
      if (
        pbDrpDwnOptns.Source.findIndex((Source) => {
          return Source.key == item.Source;
        }) == -1 &&
        item.Source
      ) {
        pbDrpDwnOptns.Source.push({
          key: item.Source,
          text: item.Source,
        });
      }
      if (
        pbDrpDwnOptns.Product.findIndex((Product) => {
          return Product.key == item.Product;
        }) == -1 &&
        item.Product
      ) {
        pbDrpDwnOptns.Product.push({
          key: item.Product,
          text: item.Product,
        });
      }
      if (
        pbDrpDwnOptns.Project.findIndex((Project) => {
          return Project.key == item.Project;
        }) == -1 &&
        item.Project
      ) {
        pbDrpDwnOptns.Project.push({
          key: item.Project,
          text: item.Project,
        });
      }
      if (Ap_AnnualPlanId) {
        if (
          pbDrpDwnOptns.Developer.findIndex((Developer) => {
            return Developer.key == item.DeveloperId;
          }) == -1 &&
          item.DeveloperId &&
          item.DeveloperEmail != "lally@goodtogreatschools.org.au"
        ) {
          let devName = allPeoples.filter(
            (dev) => dev.ID == item.DeveloperId
          )[0].text;
          pbDrpDwnOptns.Developer.push({
            key: item.DeveloperId,
            text: devName,
          });
        }
      }
    });
    if (!Ap_AnnualPlanId) {
      allPeoples.forEach((arr) => {
        if (
          pbDrpDwnOptns.Developer.findIndex((Developer) => {
            return Developer.key == arr.ID;
          }) == -1 &&
          arr.ID &&
          arr.secondaryText != "lally@goodtogreatschools.org.au" &&
          arr.secondaryText != ""
        ) {
          pbDrpDwnOptns.Developer.push({
            key: arr.ID,
            text: arr.text,
          });
        }
      });
    }
    pbDrpDwnOptns.Developer = usersOrderFunction(pbDrpDwnOptns.Developer);

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
        pbDrpDwnOptns.WeekNumber.findIndex((arr) => {
          return arr.key == i;
        }) == -1
      ) {
        pbDrpDwnOptns.WeekNumber.push({
          key: i,
          text: i.toString(),
        });
      }
    }
    for (let i = 2020; i < Pb_Year; i++) {
      if (
        pbDrpDwnOptns.Year.findIndex((arr) => {
          return arr.key == i;
        }) == -1
      ) {
        pbDrpDwnOptns.Year.push({
          key: i,
          text: i.toString(),
        });
      }
    }
    pbDrpDwnOptns.WeekNumber.sort(_sortFilterKeys);
    pbDrpDwnOptns.Year.sort(_sortFilterKeys);

    setpbDropDownOptions(pbDrpDwnOptns);
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

    if (!pbDocumentReview.Request) {
      isError = true;
      errorStatus.Request = "Please select a value for request";
    }
    if (!pbDocumentReview.Requestto) {
      isError = true;
      errorStatus.Requestto = "Please select a value for request to";
    }
    if (!pbDocumentReview.Documenttype) {
      isError = true;
      errorStatus.Documenttype = "Please select a value for document type";
    }
    if (!pbDocumentReview.Link) {
      isError = true;
      errorStatus.Link = "Please enter a value for link";
    }
    if (pbDocumentReview.Link && !pbDocumentReview.IsExternalAllow) {


      // var hasSiteOrSharePoint = /site|sharepoint/i.test(pbDocumentReview.Link);
      // if (hasSiteOrSharePoint) {
      //     var hasAspx = /aspx/i.test(pbDocumentReview.Link);
      //     isError = true;
      //     console.log("incorrect on consoele")
      //     setDocumentLinkStatus("incorrect")
      //     // return !hasAspx; // Invalid if link has "aspx"
      // }
      // setDocumentLinkStatus("correct")
      // return hasSiteOrSharePoint; // Valid if link has "site" or "sharepoint"


      const respV = isLinkValid(pbDocumentReview.Link)

      if (!respV) {
        isError = true;
        console.log("incorrect on consoele")
        setDocumentLinkStatus("incorrect")
        // return !hasAspx; // Invalid if link has "aspx"
      } else {

        setDocumentLinkStatus("correct")
      }

    }

    if (!isError) {
      setpbButtonLoader(true);
      savePBDRData();
    } else {
      setdrPBShowMessage(errorStatus);
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
  const pbValidationFunction = () => {
    let tempArronchange = pbAdhocPopup.value;
    let isError = false;

    let errorStatus = {
      BA: "",
      StartDate: "",
      EndDate: "",
      Project: "",
      Product: "",
      Title: "",
      PlannedHours: "",
    };

    if (!tempArronchange["BA"]) {
      isError = true;
      errorStatus.BA = "Please select a value for business area";
    }
    if (!tempArronchange["StartDate"]) {
      isError = true;
      errorStatus.StartDate = "Please select a value for startDate";
    }
    if (!tempArronchange["EndDate"]) {
      isError = true;
      errorStatus.EndDate = "Please select a value for endDate";
    }
    if (!tempArronchange["ProductId"]) {
      isError = true;
      errorStatus.Product = "Please select a value for product";
    }
    if (!tempArronchange["Project"]) {
      isError = true;
      errorStatus.Project = "Please select a value for name of the deliverable";
    }
    if (!tempArronchange["Title"]) {
      isError = true;
      errorStatus.Title = "Please enter a value for activity";
    }
    if (!tempArronchange["PlannedHours"]) {
      isError = true;
      errorStatus.PlannedHours = "Please enter a value for hours";
    }

    if (!isError) {
      tempArronchange["StartDate"] = moment(
        tempArronchange["StartDate"]
      ).format(DateListFormat);
      tempArronchange["EndDate"] = moment(tempArronchange["EndDate"]).format(
        DateListFormat
      );
      if (pbAdhocPopup.isNew) {
        setpbButtonLoader(true);
        setpbData(pbData.concat(tempArronchange));
        reloadFilterOptions(pbData.concat(tempArronchange));
        let pbFilter = ProductionBoardFilter(
          [...pbData.concat(tempArronchange)],
          pbFilterOptions
        );
        setpbFilterData([...pbFilter]);
        paginate(1, pbFilter);
        setpbUpdate(true);
        setpbAdhocPopup({
          visible: false,
          isNew: pbAdhocPopup.isNew,
          value: {},
        });

        //Sorting
        sortPbUpdate = true;
        sortPbFilterArr = pbFilter;
        sortPbDataArr = pbData.concat(tempArronchange);
        setpbColumns(_pbColumns);
        setpbButtonLoader(false);
      } else {
        let Index = pbData.findIndex(
          (obj) => obj.RefId == tempArronchange["RefId"]
        );
        pbData[Index] = tempArronchange;
        setpbButtonLoader(true);
        setpbData([...pbData]);
        reloadFilterOptions([...pbData]);
        let pbFilter = ProductionBoardFilter([...pbData], pbFilterOptions);

        setpbFilterData([...pbFilter]);
        paginate(1, pbFilter);
        setpbUpdate(true);
        setpbAdhocPopup({
          visible: false,
          isNew: pbAdhocPopup.isNew,
          value: {},
        });

        //Sorting
        sortPbUpdate = true;
        sortPbFilterArr = pbFilter;
        sortPbDataArr = pbData;
        setpbColumns(_pbColumns);
        setpbButtonLoader(false);
      }
    } else {
      setpbShowMessage(errorStatus);
    }
  };

  const pbDeleteItem = (id: number) => {
    sharepointWeb.lists
      .getByTitle("ProductionBoard")
      .items.getById(id)
      .delete()
      .then(() => {
        let dpData = pbMasterData.filter((arr) => {
          return arr.ID == id;
        });

        if (dpData.length > 0) {
          sharepointWeb.lists
            .getByTitle("Delivery Plan")
            .items.getById(dpData[0].DeliveryPlanID)
            .delete()
            .then(() => {
              let tempMasterArr = [...pbMasterData];
              let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
              tempMasterArr.splice(targetIndex, 1);

              let temp_ap_arr = [...pbData];
              let targetIndexapdata = temp_ap_arr.findIndex(
                (arr) => arr.ID == id
              );
              temp_ap_arr.splice(targetIndexapdata, 1);

              setpbData([...temp_ap_arr]);
              sortPbDataArr = temp_ap_arr;
              setpbMasterData([...tempMasterArr]);
              reloadFilterOptions([...tempMasterArr]);
              let pbFilter = ProductionBoardFilter(
                [...temp_ap_arr],
                pbFilterOptions
              );
              setpbFilterData(pbFilter);
              sortPbFilterArr = pbFilter;
              paginate(1, pbFilter);
              setpbDeletePopup({ condition: false, targetId: 0 });
              DeleteSuccessPopup();
            })
            .catch((error) => {
              pbErrorFunction(error, "pbDeleteItem");
            });
        } else {
          let tempMasterArr = [...pbMasterData];
          let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
          tempMasterArr.splice(targetIndex, 1);

          let temp_ap_arr = [...pbData];
          let targetIndexapdata = temp_ap_arr.findIndex((arr) => arr.ID == id);
          temp_ap_arr.splice(targetIndexapdata, 1);

          setpbData([...temp_ap_arr]);
          sortPbDataArr = temp_ap_arr;
          setpbMasterData([...tempMasterArr]);
          reloadFilterOptions([...tempMasterArr]);
          let pbFilter = ProductionBoardFilter(
            [...temp_ap_arr],
            pbFilterOptions
          );
          setpbFilterData(pbFilter);
          sortPbFilterArr = pbFilter;
          paginate(1, pbFilter);
          setpbDeletePopup({ condition: false, targetId: 0 });
          DeleteSuccessPopup();
        }
      })
      .catch((error) => {
        pbErrorFunction(error, "pbDeleteItem");
      });
  };

  const pbErrorFunction = (error, functionName: string): void => {
    console.log(error, functionName);
    let requestdata = {
      ComponentName: "Production board",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };
    Service.SPAddItem({ Listname: "Error Log", RequestJSON: requestdata }).then(
      () => {
        setpbLoader(false);
        setpbButtonLoader(false);
        setpbUpdate(!pbUpdate);
        sortPbUpdate = !pbUpdate;
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
    let tempArr = [...pbData];
    let tempDpFilterKeys = { ...pbFilterOptions };
    tempDpFilterKeys[`${key}`] = option;

    key == "WeekNumber" ? setpbWeek(option) : null;
    key == "Year" ? setpbYear(option) : null;

    // if (tempDpFilterKeys.Week == "This Week") {
    //   week = Pb_WeekNumber;
    //   year = Pb_Year;
    //   setpbWeek(Pb_WeekNumber);
    //   setpbYear(Pb_Year);
    // } else if (tempDpFilterKeys.Week == "Last Week") {
    //   week = Pb_LastWeekNumber;
    //   year = Pb_LastWeekYear;
    //   setpbWeek(Pb_LastWeekNumber);
    //   setpbYear(Pb_LastWeekYear);
    // } else if (tempDpFilterKeys.Week == "Next Week") {
    //   week = Pb_NextWeekNumber;
    //   year = Pb_NextWeekYear;
    //   setpbWeek(Pb_NextWeekNumber);
    //   setpbYear(Pb_NextWeekYear);
    // }

    if (Ap_AnnualPlanId) {
      key == "WeekNumber" || key == "Year"
        ? getCurrentPbData(
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
        setpbFilterOptions({ ...tempDpFilterKeys });
        getPbData(
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
        setpbFilterOptions({ ...tempDpFilterKeys });
        getPbData(
          loggeduserid,
          tempDpFilterKeys.WeekNumber,
          tempDpFilterKeys.Year,
          tempDpFilterKeys
        );
      }
    }

    setpbFilterOptions({ ...tempDpFilterKeys });
    let pbFilter = ProductionBoardFilter([...tempArr], tempDpFilterKeys);
    setpbFilterData(pbFilter);
    sortPbFilterArr = pbFilter;
    paginate(1, pbFilter);
  };
  const pbOnchangeItems = (RefId, key, value) => {
    let Index = pbData.findIndex((obj) => obj.RefId == RefId);
    let filIndex = pbFilterData.findIndex((obj) => obj.RefId == RefId);
    let disIndex = pbDisplayData.findIndex((obj) => obj.RefId == RefId);
    let pbBeforeData = pbData[Index];

    let pbOnchangeData = [
      {
        ID: pbBeforeData.ID,
        BA: pbBeforeData.BA,
        StartDate: pbBeforeData.StartDate,
        EndDate: pbBeforeData.EndDate,
        Source: pbBeforeData.Source,
        Project: pbBeforeData.Project,
        AnnualPlanID: pbBeforeData.AnnualPlanID,
        ProductId: pbBeforeData.ProductId,
        Product: pbBeforeData.Product,
        Title: pbBeforeData.Title,
        PlannedHours: pbBeforeData.PlannedHours,
        Monday: key == "Monday" ? value : pbBeforeData.Monday,
        Tuesday: key == "Tuesday" ? value : pbBeforeData.Tuesday,
        Wednesday: key == "Wednesday" ? value : pbBeforeData.Wednesday,
        Thursday: key == "Thursday" ? value : pbBeforeData.Thursday,
        Friday: key == "Friday" ? value : pbBeforeData.Friday,
        Saturday: key == "Saturday" ? value : pbBeforeData.Saturday,
        Sunday: key == "Sunday" ? value : pbBeforeData.Sunday,
        ActualHours: pbBeforeData.ActualHours,
        DeveloperId: pbBeforeData.DeveloperId,
        DeveloperEmail: pbBeforeData.DeveloperEmail,
        RefId: pbBeforeData.RefId,
        Week: pbBeforeData.Week,
        Year: pbBeforeData.Year,
        NotApplicable: pbBeforeData.NotApplicable,
        NotApplicableManager: pbBeforeData.NotApplicableManager,
        DeliveryPlanID: pbBeforeData.DeliveryPlanID,
        DPActualHours: pbBeforeData.DRActualHours,
        UnPlannedHours: pbBeforeData.UnPlannedHours,
        Status: pbBeforeData.Status,
        Onchange: true,
      },
    ];
    pbOnchangeData[0]["ActualHours"] =
      parseFloat(
        !isNaN(pbOnchangeData[0]["Monday"]) && pbOnchangeData[0]["Monday"]
          ? pbOnchangeData[0]["Monday"]
          : 0
      ) +
      parseFloat(
        !isNaN(pbOnchangeData[0]["Tuesday"]) && pbOnchangeData[0]["Tuesday"]
          ? pbOnchangeData[0]["Tuesday"]
          : 0
      ) +
      parseFloat(
        !isNaN(pbOnchangeData[0]["Wednesday"]) && pbOnchangeData[0]["Wednesday"]
          ? pbOnchangeData[0]["Wednesday"]
          : 0
      ) +
      parseFloat(
        !isNaN(pbOnchangeData[0]["Thursday"]) && pbOnchangeData[0]["Thursday"]
          ? pbOnchangeData[0]["Thursday"]
          : 0
      ) +
      parseFloat(
        !isNaN(pbOnchangeData[0]["Friday"]) && pbOnchangeData[0]["Friday"]
          ? pbOnchangeData[0]["Friday"]
          : 0
      ) +
      parseFloat(
        !isNaN(pbOnchangeData[0]["Saturday"]) && pbOnchangeData[0]["Saturday"]
          ? pbOnchangeData[0]["Saturday"]
          : 0
      ) +
      parseFloat(
        !isNaN(pbOnchangeData[0]["Sunday"]) && pbOnchangeData[0]["Sunday"]
          ? pbOnchangeData[0]["Sunday"]
          : 0
      );

    pbData[Index] = pbOnchangeData[0];
    pbFilterData[filIndex] = pbOnchangeData[0];
    pbDisplayData[disIndex] = pbOnchangeData[0];
    setpbData([...pbData]);
    sortPbDataArr = pbData;
    setpbFilterData([...pbFilterData]);
    sortPbFilterArr = pbFilterData;
    setpbDisplayData([...pbDisplayData]);
  };
  const drPBAddOnchange = (key, value) => {
    let tempArronchange = pbDocumentReview;
    if (key == "Request") tempArronchange.Request = value;
    else if (key == "Requestto") tempArronchange.Requestto = value;
    else if (key == "Emailcc") tempArronchange.Emailcc = value;
    else if (key == "Documenttype") tempArronchange.Documenttype = value;
    else if (key == "Link") tempArronchange.Link = value;
    else if (key == "Comments") tempArronchange.Comments = value;
    else if (key == "Confidential") tempArronchange.Confidential = value;
    else if (key == "IsExternalAllow") tempArronchange.IsExternalAllow = value;

    setpbDocumentReview(tempArronchange);
  };

  const pbAddOnchange = (key, value, PrdName) => {
    let tempArronchange = pbAdhocPopup.value;
    if (key == "BA") tempArronchange["BA"] = value;
    else if (key == "StartDate") tempArronchange["StartDate"] = value;
    else if (key == "EndDate") tempArronchange["EndDate"] = value;
    else if (key == "Product") {
      tempArronchange["ProductId"] = value;
      tempArronchange["Product"] = PrdName;
    } else if (key == "Project") tempArronchange["Project"] = value;
    else if (key == "Title") tempArronchange["Title"] = value;
    else if (key == "PlannedHours") tempArronchange["PlannedHours"] = value;
    else if (key == "UnPlannedHours") tempArronchange["UnPlannedHours"] = value;

    setpbAdhocPopup({
      visible: true,
      isNew: pbAdhocPopup.isNew,
      value: tempArronchange,
    });
  };
  const paginate = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setpbDisplayData(paginatedItems);
      setpbCurrentPage(pagenumber);
    } else {
      setpbDisplayData([]);
      setpbCurrentPage(1);
    }
  };
  const PBOnloadFilter = (data, filterValue) => {
    let tempDpFilterKeys = { ...filterValue };
    let tempArr = data.filter(
      (pb) => pb.NotApplicable != true && pb.NotApplicableManager != true
    );

    if (tempDpFilterKeys.WeekNumber) {
      tempArr = tempArr.filter((arr) => {
        // let start = moment(arr.StartDate).isoWeek();
        // let end = moment(arr.EndDate).isoWeek();
        // let today = tempDpFilterKeys.WeekNumber;
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
        let today = tempDpFilterKeys.Year.toString().concat(
          ("0" + tempDpFilterKeys.WeekNumber.toString()).slice(-2)
        );
        // moment()
        //   .year()
        //   .toString()
        //   .concat(("0" + tempDpFilterKeys.WeekNumber.toString()).slice(-2));

        return (
          parseInt(today) >= parseInt(start) && parseInt(today) <= parseInt(end)
        );
      });
    }
    if (tempDpFilterKeys.Year) {
      tempArr = tempArr.filter((arr) => {
        return arr.Year == tempDpFilterKeys.Year;
      });
    }

    // if (tempDpFilterKeys.Week == "This Week") {
    //   tempArr = tempArr.filter((arr) => {
    //     let start = moment(arr.StartDate).isoWeek();
    //     let end = moment(arr.EndDate).isoWeek();
    //     let today = Pb_WeekNumber;
    //     return today >= start && today <= end;
    //   });
    // } else if (tempDpFilterKeys.Week == "Last Week") {
    //   tempArr = tempArr.filter((arr) => {
    //     let start = moment(arr.StartDate).isoWeek();
    //     let end = moment(arr.EndDate).isoWeek();
    //     let today = Pb_LastWeekNumber;
    //     return today >= start && today <= end;
    //   });
    // } else if (tempDpFilterKeys.Week == "Next Week") {
    //   tempArr = tempArr.filter((arr) => {
    //     let start = moment(arr.StartDate).isoWeek();
    //     let end = moment(arr.EndDate).isoWeek();
    //     let today = Pb_NextWeekNumber;
    //     return today >= start && today <= end;
    //   });
    // }

    tempArr.forEach((arr, index) => {
      let dpBeforeData = tempArr[index];
      let dpOnchangeData = [
        {
          RefId: dpBeforeData.RefId,
          ID: dpBeforeData.ID,
          BA: dpBeforeData.BA,
          StartDate: dpBeforeData.StartDate,
          EndDate: dpBeforeData.EndDate,
          Source: dpBeforeData.Source,
          Project: dpBeforeData.Project,
          AnnualPlanID: dpBeforeData.AnnualPlanID,
          ProductId: dpBeforeData.ProductId,
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
          NotApplicable: dpBeforeData.NotApplicable,
          NotApplicableManager: dpBeforeData.NotApplicableManager,
          Week: dpBeforeData.Week,
          Year: dpBeforeData.Year,
          DeliveryPlanID: dpBeforeData.DeliveryPlanID,
          DPActualHours: dpBeforeData.DPActualHours,
          UnPlannedHours: dpBeforeData.UnPlannedHours,
          Status: dpBeforeData.Status,
          Onchange: true,
        },
      ];
      tempArr[index] = dpOnchangeData[0];
    });

    return tempArr;
  };
  const ProductionBoardFilter = (data, filterValue) => {
    let tempArr = data;
    let tempDpFilterKeys = { ...filterValue };

    if (tempDpFilterKeys.Showonly == "Mine") {
      tempArr = tempArr.filter((arr) => {
        return arr.DeveloperEmail == loggeduseremail;
      });
    }

    if (tempDpFilterKeys.Showonly == "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.DeveloperId == tempDpFilterKeys.Developer;
      });
    }
    if (tempDpFilterKeys.BA != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.BA == tempDpFilterKeys.BA;
      });
    }
    if (tempDpFilterKeys.Source != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Source == tempDpFilterKeys.Source;
      });
    }
    if (tempDpFilterKeys.Product != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Product == tempDpFilterKeys.Product;
      });
    }
    if (tempDpFilterKeys.Project != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Project == tempDpFilterKeys.Project;
      });
    }

    return tempArr;
  };

  const sumOfHours = () => {
    var sum: number = 0;
    // let tempArr = pbFilterData;
    let tempArr = pbFilterData.filter((arr) => {
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
    // let tempArr = pbFilterData;
    let tempArr = pbFilterData.filter((arr) => {
      return arr.UnplannedHours != true;
    });
    if (tempArr.length > 0) {
      tempArr.forEach((x) => {
        sum += parseFloat(x.ActualHours ? x.ActualHours : 0);
      });
      return sum % 1 == 0 ? sum : sum.toFixed(2);
    } else {
      return 0;
    }
  };

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _pbColumns;
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

    const newPbDataArr = _copyAndSort(
      sortPbDataArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newPbFilterArr = _copyAndSort(
      sortPbFilterArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setpbData([...newPbDataArr]);
    setpbFilterData([...newPbFilterArr]);
    paginate(1, newPbFilterArr);
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
  // Return function
  return (
    <>
      {pbLoader ? (
        <CustomLoader />
      ) : (
        <div style={{ padding: "5px 15px" }}>
          {/* {pbLoader ? <CustomLoader /> : null} */}
          <div
            className={styles.apHeaderSection}
            style={{ paddingBottom: "0" }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                marginBottom: 10,
                color: "#2392b2",
              }}
            >
              <div className={styles.dpTitle}>
                {Ap_AnnualPlanId ? (
                  <Icon
                    aria-label="ChevronLeftMed"
                    iconName="NavigateBack"
                    className={pbBigiconStyleClass.ChevronLeftMed}
                    onClick={() => {
                      pbAutoSave
                        ? alertDialogforBack()
                        : navType == "AP"
                          ? props.handleclick("AnnualPlan")
                          : props.handleclick("DeliveryPlan", Ap_AnnualPlanId);
                    }}
                  />
                ) : null}
                <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
                  Production board
                </Label>
              </div>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                marginBottom: 10,
                flexWrap: "wrap",
                color: "#2392b2",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
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
                    className={styles.toggleforAnnual}
                    onChange={(ev) => {
                      if (!Ap_AnnualPlanId) {
                        if (pbAutoSave) {
                          if (
                            confirm(
                              "You have unsaved changes, are you sure you want to leave?"
                            )
                          ) {
                            setpbChecked(!pbChecked);
                          }
                        } else {
                          setpbChecked(!pbChecked);
                        }
                      }
                    }}
                  >
                    {/* <input type="checkbox" id="toggle" /> */}
                    {pbChecked ? (
                      <input type="checkbox" id="toggle" />
                    ) : (
                      <input type="checkbox" checked id="toggle" />
                    )}
                    <span className={styles.slider}>
                      <p>Annual Plan</p>
                      <p>Activity Planner</p>
                    </span>
                  </label>
                </div>
                {!Ap_AnnualPlanId && pbWeek == Pb_WeekNumber ? (
                  <div>
                    <PrimaryButton
                      text="Ad hoc task"
                      className={pbbuttonStyleClass.buttonPrimary}
                      onClick={(_) => {
                        let adhocItem = {
                          RefId: pbData.length + 1,
                          ID: 0,
                          BA: null,
                          StartDate: new Date(),
                          EndDate: new Date(),
                          Source: "Ad hoc",
                          Project: "",
                          AnnualPlanID: null,
                          ProductId: null,
                          Product: null,
                          Title: null,
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
                          NotApplicable: false,
                          NotApplicableManager: false,
                          Week: pbWeek,
                          Year: pbYear,
                          DeliveryPlanID: null,
                          DPActualHours: 0,
                          UnPlannedHours: false,
                          Status: null,
                          Onchange: true,
                        };
                        setpbShowMessage(pbErrorStatus);
                        setpbAdhocPopup({
                          visible: true,
                          isNew: true,
                          value: adhocItem,
                        });
                      }}
                    />
                  </div>
                ) : null}
                <div className={pbProjectInfo}>
                  <Label className={pblabelStyles.titleLabel}>
                    Current week :
                  </Label>
                  <Label
                    className={pblabelStyles.labelValue}
                    style={{ maxWidth: 500 }}
                  >
                    {Pb_WeekNumber}
                  </Label>
                </div>
                <div className={pbProjectInfo}>
                  <Label className={pblabelStyles.titleLabel}>
                    Current year :
                  </Label>
                  <Label
                    className={pblabelStyles.labelValue}
                    style={{ maxWidth: 500 }}
                  >
                    {Pb_Year}
                  </Label>
                </div>
                <div className={pbProjectInfo}>
                  <Label className={pblabelStyles.titleLabel}>
                    Actual hrs/ Planned hrs :
                  </Label>
                  <Label
                    className={pblabelStyles.labelValue}
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
                  flexWrap: "wrap",
                  padding: "10px 0",
                }}
              >
                <div
                  className={pbProjectInfo}
                  style={{
                    marginRight: "20px",
                    marginTop: "-24px",
                    transform: "translateY(12px)",
                  }}
                >
                  <Label className={pblabelStyles.NORLabel}>
                    Number of records:{" "}
                    <b style={{ color: "#038387" }}>{pbFilterData.length}</b>
                  </Label>
                </div>
                {pbData.length > 0 &&
                  pbFilterOptions.Developer == loggeduserid ? (
                  <div>
                    {pbUpdate ? (
                      <div>
                        <PrimaryButton
                          iconProps={cancelIcon}
                          text="Cancel"
                          className={pbbuttonStyleClass.buttonPrimary}
                          onClick={(_) => {
                            cancelPBData();
                          }}
                        />
                        <PrimaryButton
                          iconProps={saveIcon}
                          text="Save"
                          id="pbBtnSave"
                          className={pbbuttonStyleClass.buttonSecondary}
                          onClick={(_) => {
                            setpbAutoSave(false);
                            savePBData();
                          }}
                        />
                      </div>
                    ) : (
                      <div>
                        <PrimaryButton
                          iconProps={editIcon}
                          text="Edit"
                          className={pbbuttonStyleClass.buttonPrimary}
                          onClick={() => {
                            setpbUpdate(true);
                            setpbAutoSave(true);

                            //Sorting
                            sortPbUpdate = true;
                            setpbColumns(_pbColumns);
                            setpbData(sortPbDataArr);
                            setpbFilterData(sortPbFilterArr);
                            paginate(1, sortPbFilterArr);
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
                    className={pbiconStyleClass.export}
                  />
                  Export as XLS
                </Label>
                {false ? (
                  <Icon
                    iconName="PasteAsText"
                    className={pbiconStyleClass.pblink}
                    onClick={() => {
                      // props.handleclick("ProductionBoard", Ap_AnnualPlanId);
                    }}
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
                  <Label styles={pbLabelStyles}>Business area</Label>
                  <Dropdown
                    placeholder="Select an option"
                    options={pbDropDownOptions.BA}
                    selectedKey={
                      Ap_AnnualPlanId && pbFilterData.length > 0
                        ? pbFilterData[0].BA
                        : pbFilterOptions.BA
                    }
                    styles={
                      pbFilterOptions.BA == "All"
                        ? pbDropdownStyles
                        : pbActiveDropdownStyles
                    }
                    onChange={(e, option: any) => {
                      onChangeFilter("BA", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label styles={pbLabelStyles}>Source</Label>
                  <Dropdown
                    selectedKey={pbFilterOptions.Source}
                    placeholder="Select an option"
                    options={pbDropDownOptions.Source}
                    styles={
                      pbFilterOptions.Source == "All"
                        ? pbDropdownStyles
                        : pbActiveDropdownStyles
                    }
                    onChange={(e, option: any) => {
                      onChangeFilter("Source", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label styles={pbLabelStyles}>Product</Label>
                  <Dropdown
                    selectedKey={
                      Ap_AnnualPlanId &&
                        pbFilterData.length > 0 &&
                        pbFilterData[0].Product
                        ? pbFilterData[0].Product
                        : pbFilterOptions.Product
                    }
                    placeholder="Select an option"
                    options={pbDropDownOptions.Product}
                    styles={
                      pbFilterOptions.Product == "All"
                        ? pbDropdownStyles
                        : pbActiveDropdownStyles
                    }
                    onChange={(e, option: any) => {
                      onChangeFilter("Product", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label styles={pbLabelStyles}>Project or task</Label>
                  <Dropdown
                    selectedKey={
                      Ap_AnnualPlanId && pbFilterData.length > 0
                        ? pbFilterData[0].Project
                        : pbFilterOptions.Project
                    }
                    placeholder="Select an option"
                    options={pbDropDownOptions.Project}
                    dropdownWidth={"auto"}
                    styles={
                      pbFilterOptions.Project == "All"
                        ? pbDropdownStyles
                        : pbActiveDropdownStyles
                    }
                    onChange={(e, option: any) => {
                      onChangeFilter("Project", option["key"]);
                    }}
                  />
                </div>
                <div style={{ width: "86px" }}>
                  <Label styles={pbLabelStyles}>Show only</Label>
                  <Dropdown
                    selectedKey={pbFilterOptions.Showonly}
                    placeholder="Select an option"
                    options={pbDropDownOptions.Showonly}
                    styles={showonlyDropdownActive}
                    onChange={(e, option: any) => {
                      onChangeFilter("Showonly", option["key"]);
                    }}
                  />
                </div>
                <div>
                  {/* <Label styles={pbLabelStyles}>Developer</Label> */}
                  <Dropdown
                    selectedKey={
                      pbFilterOptions.Showonly == "All"
                        ? pbFilterOptions.Developer
                        : loggeduserid
                    }
                    placeholder="Select an option"
                    options={
                      pbFilterOptions.Showonly == "Mine"
                        ? pbDropDownOptions.DeveloperMine
                        : pbDropDownOptions.Developer
                    }
                    styles={pbActiveDropdownStyles}
                    style={{ marginTop: 25 }}
                    onChange={(e, option: any) => {
                      onChangeFilter("Developer", option["key"]);
                    }}
                  />
                </div>
                {/* <div>
                  <Label styles={pbLabelStyles}>Week</Label>
                  <Dropdown
                    selectedKey={pbFilterOptions.Week}
                    placeholder="Select an option"
                    options={pbDropDownOptions.Week}
                    styles={pbActiveDropdownStyles}
                    onChange={(e, option: any) => {
                      onChangeFilter("Week", option["key"]);
                    }}
                  />
                </div> */}
                <div>
                  <Label styles={pbShortLabelStyles}>Week</Label>
                  <Dropdown
                    selectedKey={pbFilterOptions.WeekNumber}
                    placeholder="Select an option"
                    options={pbDropDownOptions.WeekNumber}
                    styles={pbActiveShortDropdownStyles}
                    onChange={(e, option: any) => {
                      onChangeFilter("WeekNumber", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label styles={pbShortLabelStyles}>Year</Label>
                  <Dropdown
                    selectedKey={pbFilterOptions.Year}
                    placeholder="Select an option"
                    options={pbDropDownOptions.Year}
                    styles={pbActiveShortDropdownStyles}
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
                      className={pbiconStyleClass.refresh}
                      onClick={() => {
                        if (pbAutoSave) {
                          if (
                            confirm(
                              "You have unsaved changes, are you sure you want to leave?"
                            )
                          ) {
                            setpbWeek(Pb_WeekNumber);
                            setpbYear(Pb_Year);
                            setpbFilterOptions({ ...pbFilterKeys });

                            if (Ap_AnnualPlanId) {
                              setpbData([...pbMasterData]);
                              sortPbDataArr = pbMasterData;
                              let pbFilter = ProductionBoardFilter(
                                [...pbMasterData],
                                pbFilterKeys
                              );
                              setpbFilterData(pbFilter);
                              sortPbFilterArr = pbFilter;
                              paginate(1, pbFilter);
                              setpbUpdate(false);
                              sortPbUpdate = false;

                              setpbColumns(_pbColumns);
                              getCurrentPbData(
                                Pb_WeekNumber,
                                Pb_Year,
                                pbFilterKeys
                              );
                            } else {
                              setpbUpdate(false);
                              sortPbUpdate = false;
                              setpbColumns(_pbColumns);
                              getPbData(
                                loggeduserid,
                                Pb_WeekNumber,
                                Pb_Year,
                                pbFilterKeys
                              );
                            }
                          }
                        } else {
                          setpbWeek(Pb_WeekNumber);
                          setpbYear(Pb_Year);
                          setpbFilterOptions({ ...pbFilterKeys });

                          if (Ap_AnnualPlanId) {
                            setpbData([...pbMasterData]);
                            sortPbDataArr = pbMasterData;
                            let pbFilter = ProductionBoardFilter(
                              [...pbMasterData],
                              pbFilterKeys
                            );
                            setpbFilterData(pbFilter);
                            sortPbFilterArr = pbFilter;
                            paginate(1, pbFilter);
                            setpbUpdate(false);
                            sortPbUpdate = false;

                            setpbColumns(_pbColumns);
                            getCurrentPbData(
                              Pb_WeekNumber,
                              Pb_Year,
                              pbFilterKeys
                            );
                          } else {
                            setpbUpdate(false);
                            sortPbUpdate = false;

                            setpbColumns(_pbColumns);
                            getPbData(
                              loggeduserid,
                              Pb_WeekNumber,
                              Pb_Year,
                              pbFilterKeys
                            );
                          }
                        }
                      }}
                    />
                  </div>
                </div>
              </div>

              {/* <div
            className={pbProjectInfo}
            style={{ marginLeft: "20px", transform: "translateY(12px)" }}
          >
            <Label className={pblabelStyles.NORLabel}>
              Number of records:{" "}
              <b style={{ color: "#038387" }}>{pbFilterData.length}</b>
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
          {pbChecked ? (
            <div style={{ marginTop: "10px" }}>
              <DetailsList
                items={pbDisplayData}
                columns={sortPbUpdate ? _pbColumns : pbColumns}
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
                {pbFilterData.length > 0 ? (
                  <Pagination
                    currentPage={pbcurrentPage}
                    totalPages={
                      pbFilterData.length > 0
                        ? Math.ceil(pbFilterData.length / totalPageItems)
                        : 1
                    }
                    onChange={(page) => {
                      paginate(page, pbFilterData);
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
              "ActivityProductionBoard",
              pbSwitchID,
              pbSwitchType,
              Ap_AnnualPlanId ? Ap_AnnualPlanId + "-" + navType : null
            )
          )}

          <Modal isOpen={pbModalBoxVisibility} isBlocking={false}>
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
                    errorMessage={drPBShowMessage.Request}
                    label="Request"
                    placeholder="Select an option"
                    options={pbModalBoxDropDownOptions.Request}
                    styles={pbModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      drPBAddOnchange("Request", option["key"]);
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
                    className={pbModalBoxPP}
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
                        ? drPBAddOnchange("Requestto", selectedUser[0]["ID"])
                        : drPBAddOnchange("Requestto", "");
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
                    {drPBShowMessage.Requestto}
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
                    className={pbModalBoxPP}
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={5}
                    // styles={{
                    //   root: {
                    //     width: 300,
                    //     margin: "10px 20px",
                    //     selectors: {
                    //       ".ms-BasePicker-text": {
                    //         height: 36,
                    //         padding: "3px 10px",
                    //         border: "1px solid black",
                    //         borderRadius: 4,
                    //       },
                    //     },
                    //     ".ms-Persona-primaryText": { fontWeight: 600 },
                    //   },
                    // }}
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
                        ? drPBAddOnchange("Emailcc", selectedId)
                        : drPBAddOnchange("Emailcc", "");
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
                    placeholder="Add new name of the deliverable"
                    defaultValue={pbDocumentReview.Project}
                    disabled={true}
                    styles={pbTxtBoxStyles}
                    className={styles.projectField}
                    onChange={(e, value: string) => { }}
                  />
                </div>
                <div>
                  <Dropdown
                    label="Document type"
                    required={true}
                    errorMessage={drPBShowMessage.Documenttype}
                    placeholder="Select an option"
                    options={pbModalBoxDropDownOptions.Documenttype}
                    styles={pbModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      drPBAddOnchange("Documenttype", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <TextField
                    label="Link"
                    placeholder="Add link"
                    errorMessage={drPBShowMessage.Link}
                    required={true}
                    styles={pbTxtBoxStyles}
                    onChange={(e, value: string) => {
                      drPBAddOnchange("Link", value);
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
                    styles={pbMultiTxtBoxStyles}
                    onChange={(e, value: string) => {
                      drPBAddOnchange("Comments", value);
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
                      drPBAddOnchange(
                        "Confidential",
                        !pbDocumentReview.Confidential
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
                      drPBAddOnchange(
                        "IsExternalAllow",
                        !pbDocumentReview.IsExternalAllow
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
                  {pbButtonLoader ? (
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
                    setpbModalBoxVisibility(false);
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
          <Modal isOpen={pbAdhocPopup.visible} isBlocking={false}>
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
                    label="Business area"
                    placeholder="Select a business area"
                    required={true}
                    options={pbModalBoxDropDownOptions.BA}
                    errorMessage={pbShowMessage.BA}
                    selectedKey={pbAdhocPopup.value["BA"]}
                    styles={pbModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      pbAddOnchange("BA", option["key"], null);
                    }}
                  />
                </div>
                <div>
                  <DatePicker
                    label="Start date"
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    styles={pbModalBoxDatePickerStyles}
                    formatDate={dateFormater}
                    value={pbAdhocPopup.value["StartDate"]}
                    onSelectDate={(value: any) => {
                      pbAddOnchange("StartDate", value, null);
                    }}
                  />
                </div>
                <div>
                  <DatePicker
                    label="End date"
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    styles={pbModalBoxDatePickerStyles}
                    formatDate={dateFormater}
                    value={pbAdhocPopup.value["EndDate"]}
                    onSelectDate={(value: any) => {
                      pbAddOnchange("EndDate", value, null);
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
                    options={pbModalBoxDropDownOptions.Project}
                    errorMessage={pbShowMessage.Project}
                    selectedKey={pbAdhocPopup.value["Project"]}
                    styles={pbModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      pbAddOnchange("Project", option["key"], null);
                    }}
                  />
                </div>
                <div>
                  <Dropdown
                    label="Product"
                    required={true}
                    placeholder="Select a product"
                    options={pbModalBoxDropDownOptions.Product}
                    errorMessage={pbShowMessage.Product}
                    selectedKey={pbAdhocPopup.value["ProductId"]}
                    styles={pbModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      pbAddOnchange("Product", option["key"], option["text"]);
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
                    label="Activity"
                    placeholder="Add activity"
                    errorMessage={pbShowMessage.Title}
                    value={pbAdhocPopup.value["Title"]}
                    required={true}
                    styles={pbTxtBoxStyles}
                    className={styles.projectField}
                    onChange={(e, value: string) => {
                      pbAddOnchange("Title", value, null);
                    }}
                  />
                </div>
                <div>
                  <TextField
                    label="Hours"
                    placeholder="Add hours"
                    errorMessage={pbShowMessage.PlannedHours}
                    value={pbAdhocPopup.value["PlannedHours"]}
                    required={true}
                    styles={pbTxtBoxStyles}
                    className={styles.projectField}
                    onChange={(e, value: string) => {
                      parseFloat(value)
                        ? pbAddOnchange("PlannedHours", value, null)
                        : pbAddOnchange("PlannedHours", null, null);
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
                          width: "300px",
                        }}
                      >
                        Unplanned hours
                      </div>
                    }
                    inlineLabel
                    checked={pbAdhocPopup.value["UnPlannedHours"]}
                    style={{ transform: "translateX(100px)", marginLeft: 25 }}
                    onChange={(ev) => {
                      pbAddOnchange(
                        "UnPlannedHours",
                        !pbAdhocPopup.value["UnPlannedHours"],
                        null
                      );
                    }}
                  />
                </div>
              </div>
              <div className={styles.apModalBoxButtonSection}>
                <button
                  className={styles.apModalBoxSubmitBtn}
                  onClick={(_) => {
                    pbValidationFunction();
                  }}
                  style={{ display: "flex" }}
                >
                  {pbButtonLoader ? (
                    <Spinner />
                  ) : (
                    <span>
                      <Icon
                        iconName="Save"
                        style={{ position: "relative", top: 3, left: -8 }}
                      />
                      {pbAdhocPopup.isNew ? "Submit" : "Update"}
                    </span>
                  )}
                </button>
                <button
                  className={styles.apModalBoxBackBtn}
                  onClick={(_) => {
                    setpbAdhocPopup({
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
            <Modal isOpen={pbDeletePopup.condition} isBlocking={true}>
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
                    setpbButtonLoader(true);
                    pbDeleteItem(pbDeletePopup.targetId);
                  }}
                  className={styles.apDeletePopupYesBtn}
                >
                  {pbButtonLoader ? <Spinner /> : "Yes"}
                </button>
                <button
                  onClick={(_) => {
                    setpbDeletePopup({ condition: false, targetId: 0 });
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

export default ProductionBoard;
