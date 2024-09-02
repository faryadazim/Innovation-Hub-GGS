import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import {
  DetailsList,
  DetailsListLayoutMode,
  IDetailsListStyles,
  SelectionMode,
  Icon,
  Label,
  ILabelStyles,
  Dropdown,
  IDropdownStyles,
  NormalPeoplePicker,
  Persona,
  PersonaPresence,
  PersonaSize,
  DatePicker,
  Spinner,
  PrimaryButton,
  SearchBox,
  ISearchBoxStyles,
  TooltipHost,
  TooltipOverflowMode,
  TextField,
  Checkbox,
  Modal,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import { FontIcon, IIconProps } from "@fluentui/react/lib/Icon";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { maxBy } from "lodash";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";

const saveIcon: IIconProps = { iconName: "Save" };
const editIcon: IIconProps = { iconName: "Edit" };
const cancelIcon: IIconProps = { iconName: "Cancel" };

let DateListFormat = "DD/MM/YYYY";
let DatePickerFormat = "YYYY-MM-DDT14:00:00Z";

const ActivityDeliveryPlan = (props: any) => {
  // Variable-Declaration-Section Starts
  const sharepointWeb = Web(props.URL);
  const activityPlan_ID = props.ActivityPlanID;

  const activityPlanListName = "Activity Plan";
  const adpListName = "Activity Delivery Plan";
  const templateListName = "Activity Delivery Plan Template";
  const activityPBListName = "ActivityProductionBoard";

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  const adpCurrentWeekNumber = moment().isoWeek();
  const adpCurrentYear = moment().year();

  // const adpAllitems = [];
  const allPeoples = props.peopleList;
  const _adpColumns = [
    {
      key: "Lesson",
      name: "Section",
      fieldName: "Lesson",
      minWidth: 80,
      maxWidth: 100,

      onRender: (item, index) =>
        index == 0 ? (
          <>
            <input
              id="forFocus"
              type="text"
              style={{
                width: 0,
                height: 0,
                border: "none",
                position: "absolute",
                top: 0,
                left: 0,
                padding: 0,
              }}
            />
            <TooltipHost
              id={item.ID}
              content={item.Lesson}
              overflowMode={TooltipOverflowMode.Parent}
            >
              <span aria-describedby={item.ID}>{item.Lesson}</span>
            </TooltipHost>
          </>
        ) : (
          <>
            <TooltipHost
              id={item.ID}
              content={item.Lesson}
              overflowMode={TooltipOverflowMode.Parent}
            >
              <span aria-describedby={item.ID}>{item.Lesson}</span>
            </TooltipHost>
          </>
        ),
    },
    {
      key: "Steps",
      name: "Steps",
      fieldName: "Steps",
      minWidth: 80,
      maxWidth: 250,

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
      key: "Complete",
      name: "IsComplete",
      fieldName: "IsCompleteStatus",
      minWidth: 85,
      maxWidth: 85,
      onRender: (item, Index) => (
        <Checkbox
          styles={{ root: { marginTop: 3 } }}
          data-id={item.ID}
          disabled={!adpEditFlag || item.Status == "Completed" ? true : false}
          checked={item.IsCompleteStatus}
          onChange={(ev) => {
            adpActivityResponseHandler(
              item.OrderId,
              "IsCompleteStatus",
              ev.target["checked"]
            );
          }}
        />
      ),
    },
    {
      key: "PH",
      name: "PH",
      fieldName: "PH",
      minWidth: 80,
      maxWidth: 100,

      onRender: (item, index: number) => {
        if (adpEditFlag && item.MinPH && item.MinPH) {
          return (
            <TextField
              styles={{
                root: {
                  selectors: {
                    ".ms-TextField-fieldGroup": {
                      borderRadius: 4,
                      border: "1px solid",
                      height: 28,
                      width: 80,
                      input: {
                        borderRadius: 4,
                      },
                      color: item.PHError ? "#d0342c" : "#000",
                    },
                    // ".ms-TextField-field": {
                    //   color: item.PHError ? "#d0342c" : "#000",
                    // },
                  },
                },
              }}
              data-id={item.ID}
              value={item.PH}
              placeholder={`${item.MinPH} - ${item.MaxPH}hr`}
              onChange={(e: any) => {
                let valPH = e.target.value;
                if (parseFloat(valPH)) {
                  adpActivityResponseHandler(item.OrderId, "PH", valPH);
                } else {
                  adpActivityResponseHandler(item.OrderId, "PH", null);
                }
              }}
            />
          );
        } else if (item.PHWeek) {
          let valPH = item.PH.toString();
          valPH = valPH.split(".");
          let resultPH;
          if (valPH.length > 1) {
            if (valPH[0] == "0") {
              resultPH = Math.round((item.PH - valPH[0]) * 7) + " D ";
            } else {
              resultPH =
                Math.round(valPH[0]) +
                " W " +
                Math.round((item.PH - valPH[0]) * 7) +
                " D ";
            }
          } else {
            resultPH = Math.round(item.PH) + "W";
          }
          return resultPH;
        } else {
          return item.PH;
        }
      },
    },
    {
      key: "Start",
      name: "Start date",
      fieldName: "Start",
      minWidth: 100,
      maxWidth: 120,

      onRender: (item, index: number) =>
        adpEditFlag ? (
          <>
            <DatePicker
              placeholder="Select a start date"
              formatDate={dateFormater}
              // minDate={new Date(item.Start)}
              // maxDate={new Date(item.End)}
              styles={{
                textField: {
                  transform: "translateY(3px)",
                  selectors: {
                    ".ms-TextField-fieldGroup": {
                      borderColor: item.dateError ? "#d0342c" : "#000",
                      borderRadius: 4,
                      border: "1px solid",
                      height: 23,
                      input: {
                        borderRadius: 4,
                      },
                    },
                    ".ms-TextField-field": {
                      color: item.dateError ? "#d0342c" : "#000",
                    },
                    ".ms-DatePicker-event--without-label": {
                      color: item.dateError ? "#d0342c" : "#000",
                      paddingTop: 3,
                    },
                  },
                },
                readOnlyTextField: {
                  lineHeight: 22,
                },
              }}
              value={
                item.Start
                  ? new Date(
                      moment(item.Start, DateListFormat).format(
                        DatePickerFormat
                      )
                    )
                  : new Date()
              }
              onSelectDate={(value: any) => {
                adpActivityResponseHandler(item.OrderId, "Start", value);
              }}
            />
          </>
        ) : (
          <>
            {item.Start ? (
              <>{item.Start}</>
            ) : (
              <>{moment().format(DateListFormat)}</>
            )}
          </>
        ),
    },
    {
      key: "End",
      name: "End date",
      fieldName: "End",
      minWidth: 100,
      maxWidth: 120,

      onRender: (item, index: number) =>
        adpEditFlag ? (
          <>
            <DatePicker
              placeholder="Select a end date"
              formatDate={dateFormater}
              // minDate={new Date(item.Start)}
              // maxDate={new Date(item.End)}
              styles={{
                textField: {
                  transform: "translateY(3px)",
                  selectors: {
                    ".ms-TextField-fieldGroup": {
                      borderColor: item.dateError ? "#d0342c" : "#000",
                      borderRadius: 4,
                      border: "1px solid",
                      height: 23,
                      input: {
                        borderRadius: 4,
                      },
                    },
                    ".ms-TextField-field": {
                      color: item.dateError ? "#d0342c" : "#000",
                    },
                    ".ms-DatePicker-event--without-label": {
                      color: item.dateError ? "#d0342c" : "#000",
                      paddingTop: 3,
                    },
                  },
                },
                readOnlyTextField: {
                  lineHeight: 22,
                },
              }}
              value={
                item.End
                  ? new Date(
                      moment(item.End, DateListFormat).format(DatePickerFormat)
                    )
                  : new Date()
              }
              onSelectDate={(value: any) => {
                adpActivityResponseHandler(item.OrderId, "End", value);
              }}
            />
          </>
        ) : (
          <>
            {item.End ? (
              <>{item.End}</>
            ) : (
              <>{moment().format(DateListFormat)}</>
            )}
          </>
        ),
    },
    {
      key: "Status",
      name: "Status",
      fieldName: "Status",
      minWidth: 80,
      maxWidth: 120,

      onRender: (item) => (
        <>
          {item.Status == "Completed" ? (
            <div className={adpStatusStyleClass.completed}>{item.Status}</div>
          ) : item.Status == "Scheduled" ? (
            <div className={adpStatusStyleClass.scheduled}>{item.Status}</div>
          ) : item.Status == "On schedule" ? (
            <div className={adpStatusStyleClass.onSchedule}>{item.Status}</div>
          ) : item.Status == "Behind schedule" ? (
            <div className={adpStatusStyleClass.behindescheduled}>
              {item.Status}
            </div>
          ) : item.Status == "On hold" ? (
            <div className={adpStatusStyleClass.Onhold}>{item.Status}</div>
          ) : (
            ""
          )}
        </>
      ),
    },
    {
      key: "Developer",
      name: "Developer",
      fieldName: "Developer",
      minWidth: 150,
      maxWidth: 200,

      onRender: (item) =>
        adpEditFlag ? (
          <>
            <NormalPeoplePicker
              styles={{
                root: {
                  selectors: {
                    ".ms-SelectionZone": {
                      height: 24,
                    },
                    ".ms-BasePicker-text": {
                      height: 24,
                      padding: 1,
                      border: "1px solid #000",
                      borderRadius: 4,
                      marginTop: -6,
                      marginRight: 20,
                    },
                  },
                },
              }}
              onResolveSuggestions={GetUserDetails}
              itemLimit={1}
              selectedItems={allPeoples.filter((people) => {
                return (
                  people.ID == (item.Developer.id ? item.Developer.id : null)
                );
              })}
              onChange={(selectedUser) => {
                adpActivityResponseHandler(
                  item.OrderId,
                  "Developer",
                  selectedUser[0] ? selectedUser[0]["ID"] : null
                );
              }}
            />
          </>
        ) : (
          <>
            {newDataFlag ? (
              <div style={{ display: "flex" }}>
                <div
                  style={{
                    marginTop: "-6px",
                  }}
                  title={item.Developer ? item.Developer.name : "no Name"}
                >
                  <Persona
                    size={PersonaSize.size32}
                    presence={PersonaPresence.none}
                    imageUrl={
                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                      `${item.Developer.email}`
                    }
                  />
                </div>
                <div>
                  <span style={{ fontSize: "13px" }}>
                    {item.Developer.name}
                  </span>
                </div>
              </div>
            ) : item.Developer.id ? (
              <div style={{ display: "flex" }}>
                <div
                  style={{
                    marginTop: "-6px",
                  }}
                  title={item.Developer ? item.Developer.name : "no Name"}
                >
                  <Persona
                    size={PersonaSize.size32}
                    presence={PersonaPresence.none}
                    imageUrl={
                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                      `${item.Developer.email}`
                    }
                  />
                </div>
                <div>
                  <span style={{ fontSize: "13px" }}>
                    {item.Developer.name}
                  </span>
                </div>
              </div>
            ) : (
              ""
            )}
          </>
        ),
    },
  ];
  const adpStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
  });
  const adpStatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      adpStatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      adpStatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      adpStatusStyle,
    ],
    behindescheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      adpStatusStyle,
    ],
    Onhold: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      adpStatusStyle,
    ],
  });
  const gridStyles: Partial<IDetailsListStyles> = {
    root: {
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          ".ms-DetailsRow-fields": {
            alignItems: "center",
            height: 38,
          },
          ".ms-GroupHeader-title ": {
            fontSize: 14,
            color: "#000",
          },
          ".ms-GroupHeader > div": {
            height: 43,
            // backgroundColor: "#03838752 !important",
            borderBottom: "2px solid #eee",
          },
          ".ms-GroupHeader-expand:hover": {
            backgroundColor: "transparent",
          },
        },
      },
    },

    headerWrapper: {
      flex: "0 0 auto",
    },
    contentWrapper: {
      // flex: "1 1 auto",
      // overflow: "hidden",
      height: "calc(100vh - 310px)",
      overflowX: "hidden",
      overflowY: "auto",
    },
  };
  const noDatagridStyles: Partial<IDetailsListStyles> = {
    root: {
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          ".ms-DetailsRow-fields": {
            alignItems: "center",
            height: 38,
          },
          ".ms-GroupHeader-title ": {
            fontSize: 14,
            color: "#000",
          },
          ".ms-GroupHeader > div": {
            height: 43,
            // backgroundColor: "#03838752 !important",
            borderBottom: "2px solid #eee",
          },
          ".ms-GroupHeader-expand:hover": {
            backgroundColor: "transparent",
          },
        },
      },
    },

    headerWrapper: {
      flex: "0 0 auto",
    },
    contentWrapper: {
      // flex: "1 1 auto",
      // overflow: "hidden",
      // height: "calc(100vh - 310px)",
      overflowX: "hidden",
      overflowY: "auto",
    },
  };
  const adpDrpDwnOptns = {
    developerOptns: [{ key: "All", text: "All" }],
    stepsOptns: [{ key: "All", text: "All" }],
    lessonOptns: [{ key: "All", text: "All" }],
    statusOptns: [{ key: "All", text: "All" }],
    weekOptns: [{ key: "All", text: "All" }],
    yearOptns: [{ key: "All", text: "All" }],
  };
  const adpFilterKeys = {
    developer: "All",
    step: "All",
    lesson: "",
    status: "All",
    week: "All",
    year: "All",
  };
  // const adpFilterKeys = { developer: "All", step: "All", lesson: "All" };

  // Variable-Declaration-Section Ends
  // Styles-Section Starts
  const adpLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 165,
      marginTop: 5,
      marginRight: 10,
      fontSize: 13,
      color: "#323130",
    },
  };
  const adpDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 200,
      marginTop: 5,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "1px solid #E8E8EA",
      borderRadius: 4,
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
  const adpActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 200,
      marginTop: 5,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
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
    caretDown: { fontSize: 14, color: "#038387", fontWeight: 600 },
  };

  const adpShortDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 75,
      marginTop: 5,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "1px solid #E8E8EA",
      borderRadius: 4,
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
  const adpActiveShortDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 75,
      marginTop: 5,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
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
    caretDown: { fontSize: 14, color: "#038387", fontWeight: 600 },
    callout: {
      maxHeight: 300,
    },
  };

  const adpSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 200,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
      marginTop: "3px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const adpActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 200,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "2px solid #038387",
      borderRadius: "4px",
      marginTop: "3px",
    },
    field: { fontWeight: 600, color: "#038387" },
    icon: { fontSize: 14, color: "#038387" },
  };
  const adpCommonStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: 25,
    fontWeight: "600",
    padding: 3,
    width: 100,
    display: "flex",
    justifyContent: "center",
  });
  const adpStatusStyles = mergeStyleSets({
    completed: [
      {
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      adpCommonStatusStyle,
    ],
    scheduled: [
      {
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      adpCommonStatusStyle,
    ],
    onSchedule: [
      {
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      adpCommonStatusStyle,
    ],
    behindescheduled: [
      {
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      adpCommonStatusStyle,
    ],
  });
  const adpbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const adpbuttonStyleClass = mergeStyleSets({
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
      adpbuttonStyle,
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
      adpbuttonStyle,
    ],
  });
  const adpIconStyleClass = mergeStyleSets({
    navArrow: [
      {
        cursor: "pointer",
        color: "#2392b2",
        fontSize: 24,
        marginTop: "3px",
        marginRight: 12,
      },
    ],
    navArrowDisabled: [
      {
        cursor: "pointer",
        color: "#ababab",
        fontSize: 24,
        marginTop: "3px",
        marginRight: 12,
      },
    ],
    link: [
      {
        fontSize: 17,
        height: 16,
        width: 16,
        color: "#fff",
        backgroundColor: "#038387",
        cursor: "pointer",
        padding: 8,
        borderRadius: 3,
        marginLeft: 10,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    linkDisabled: [
      {
        fontSize: 18,
        height: 16,
        width: 19,
        color: "#fff",
        backgroundColor: "#ababab",
        cursor: "not-allowed",
        padding: 8,
        borderRadius: 3,
        marginLeft: 10,
      },
    ],
    refresh: [
      {
        fontSize: 18,
        height: 16,
        width: 19,
        color: "#fff",
        backgroundColor: "#038387",
        cursor: "pointer",
        padding: 8,
        borderRadius: 3,
        marginTop: 34,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    save: [
      {
        fontSize: "18px",
        color: "#fff",
        paddingRight: 10,
      },
    ],
    edit: [
      {
        fontSize: "18px",
        color: "#fff",
        paddingRight: 10,
      },
    ],
    export: [
      {
        color: "black",
        fontSize: "18px",
        height: 20,
        width: 20,
        cursor: "pointer",
      },
    ],
  });
  const adpCommonStyles = mergeStyleSets({
    titleLabel: {
      color: "#2392B2 !important",
      fontWeight: "500",
      fontSize: 17,
    },
    inputLabel: {
      color: "#2392B2 !important",
      display: "block",
      fontWeight: "500",
      margin: "5px 0",
    },
    inputValue: {
      color: "#000",
      fontWeight: "500",
      fontSize: 13,
    },
    inputField: {
      margin: "10px 0",
    },
    dateGridValidationErrorLabel: {
      color: "#d0342c !important",
      fontWeight: 600,
      marginLeft: 20,
    },
  });
  // Styles-Section Ends
  // States-Declaration Starts
  const [adpReRender, setAdpReRender] = useState(true);
  const [currentUser, setCurrentUser] = useState({});
  const [activtyPlanItem, setActivtyPlanItem] = useState([]);
  const [activityPB, setActivityPB] = useState([]);
  const [group, setgroup] = useState([]);
  const [adpMasterData, setAdpMasterData] = useState([]);
  const [adpData, setAdpData] = useState([]);
  const [adpDropDownOptions, setAdpDropDownOptions] = useState(adpDrpDwnOptns);
  const [adpFilters, setAdpFilters] = useState(adpFilterKeys);
  const [adpActivityResponseData, setAdpActivityResponseData] = useState([]);
  const [adpEditFlag, setAdpEditFlag] = useState(false);
  const [newDataFlag, setNewDataFlag] = useState(false);
  const [adpItemAddFlag, setAdpItemAddFlag] = useState(false);
  const [adpLoader, setAdpLoader] = useState("noLoader");
  const [adpAutoSave, setAdpAutoSave] = useState(false);

  const [adpSDSort, setAdpSDSort] = useState("");
  const [adpEDSort, setAdpEDSort] = useState("");

  const [AdpIsCompleted, setAdpIsCompleted] = useState(false);

  const [AdpConfirmationPopup, setAdpConfirmationPopup] = useState({
    condition: false,
    isNew: false,
  });

  window.onbeforeunload = function (e) {
    if (adpAutoSave) {
      let dialogText =
        "You have unsaved changes, are you sure you want to leave?";
      e.returnValue = dialogText;
      return dialogText;
    }
  };

  // States-Declaration Ends
  //Function-Section Starts
  const generateExcel = () => {
    let arrExport = adpActivityResponseData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Lesson", key: "Lesson", width: 25 },
      { header: "Steps", key: "Steps", width: 25 },
      { header: "PH", key: "PH", width: 25 },
      { header: "Start date", key: "Start", width: 25 },
      { header: "End date", key: "End", width: 25 },
      { header: "Status", key: "Status", width: 30 },
      { header: "Developer", key: "Developer", width: 60 },
    ];
    arrExport.forEach((item) => {
      worksheet.addRow({
        Lesson: item.Lesson ? item.Lesson : null,
        Steps: item.Steps ? item.Steps : null,
        PH: item.PH ? item.PH : "",
        Start: item.Start ? item.Start : "",
        End: item.End ? item.End : "",
        Status: item.Status ? item.Status : "",
        Developer: item.Developer ? item.Developer.name : null,
      });
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1"].map((key) => {
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
          `ActivityPlan-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };
  const adpGetCurrentUserDetails = () => {
    sharepointWeb.currentUser
      .get()
      .then((user) => {
        let adpCurrentUser = {
          Name: user.Title,
          Email: user.Email,
          Id: user.Id,
        };
        setCurrentUser({ ...adpCurrentUser });
      })
      .catch((err) => {
        adpErrorFunction(err, "adpGetCurrentUserDetails");
      });
  };
  const getActivityPlanItem = () => {
    let _adpItem = [];

    sharepointWeb.lists
      .getByTitle(activityPlanListName)
      .items.getById(activityPlan_ID)
      .get()
      .then((item) => {
        _adpItem.push({
          ID: item.Id ? item.Id : "",
          Lesson: item.Lessons ? item.Lessons : "",
          Project: item.Project ? item.Project : "",
          Product: item.Product ? item.Product : "",
          ProjectVersion: item.ProjectVersion ? item.ProjectVersion : "V1",
          ProductVersion: item.ProductVersion ? item.ProductVersion : "V1",
          Types: item.Types ? item.Types : "",
          Status: item.Status ? item.Status : null,
        });

        let _adpLessons = [];
        let lessons = _adpItem[0].Lesson.split(";");

        lessons.forEach((ls) => {
          console.log(ls.split("~")[4]);
          _adpLessons.push({
            ID: ls.split("~")[0],
            Name: ls.split("~")[1],
            StartDate: ls.split("~")[2],
            EndDate: ls.split("~")[3],
            DeveloperId:
              ls.split("~")[4] != "NaN" && ls.split("~")[4] != null
                ? parseInt(ls.split("~")[4])
                : null,
            DeveloperName:
              ls.split("~")[4] != "NaN" &&
              ls.split("~")[4] != "null" &&
              allPeoples.length > 0 &&
              allPeoples.filter((ap) => {
                return ap.ID == ls.split("~")[4];
              }).length > 0
                ? allPeoples.filter((ap) => {
                    return ap.ID == ls.split("~")[4];
                  })[0].text
                : null,
            DeveloperEmail:
              ls.split("~")[4] != "NaN" &&
              ls.split("~")[4] != "null" &&
              allPeoples.length > 0 &&
              allPeoples.filter((ap) => {
                return ap.ID == ls.split("~")[4];
              }).length > 0
                ? allPeoples.filter((ap) => {
                    return ap.ID == ls.split("~")[4];
                  })[0].secondaryText
                : null,
          });
        });
        adpGetData(_adpItem[0], _adpLessons);
        setActivtyPlanItem([..._adpItem]);
      })
      .catch((err) => {
        adpErrorFunction(err, "getActivityPlanItem");
      });
  };
  const getActivityPBData = () => {
    sharepointWeb.lists
      .getByTitle(activityPBListName)
      .items.filter(
        `ActivityPlanID eq '${activityPlan_ID}' 
        and Week eq '${adpCurrentWeekNumber}' 
        and Year eq '${adpCurrentYear}'`
      )
      .top(5000)
      .get()
      .then((items) => {
        setActivityPB([...items]);
      })
      .catch((err) => {
        adpErrorFunction(err, "getActivityPBData");
      });
  };
  const adpGetData = (adpItem: any, lessons) => {
    let adpAllitems = [];
    sharepointWeb.lists
      .getByTitle(adpListName)
      .items.select(
        "*",
        "Developer/Title",
        "Developer/Id",
        "Developer/EMail",
        "FieldValuesAsText/StartDate",
        "FieldValuesAsText/EndDate"
      )
      .expand("Developer,FieldValuesAsText")
      .filter(`ActivityPlanID eq ${activityPlan_ID}`)
      .orderBy("OrderId", true)
      .top(5000)
      .get()
      .then((items) => {
        console.log(items);
        if (items.length > 0) {
          items.forEach((item, index) => {
            adpAllitems.push({
              OrderId: index,
              LessonID: item.LessonID,
              ID: item.Id ? item.Id : "",
              Steps: item.Title ? item.Title : "",
              PH: item.PlannedHours ? item.PlannedHours : "",
              MinPH: item.MinPH ? item.MinPH : "",
              MaxPH: item.MaxPH ? item.MaxPH : "",
              Project: item.Project ? item.Project : "",
              Lesson: item.Lesson ? item.Lesson : "",
              Start: item.StartDate
                ? moment(
                    item["FieldValuesAsText"].StartDate,
                    DateListFormat
                  ).format(DateListFormat)
                : null,
              End: item.EndDate
                ? moment(
                    item["FieldValuesAsText"].EndDate,
                    DateListFormat
                  ).format(DateListFormat)
                : null,
              Developer: item.DeveloperId
                ? {
                    name: item.Developer.Title,
                    id: item.Developer.Id,
                    email: item.Developer.EMail,
                  }
                : "",
              Status: item.Status ? item.Status : "noData",
              IsCompleteStatus: item.Status == "Completed" ? true : false,
              IsCompleteNew: false,
              AH: item.ActualHours ? item.ActualHours : 0,
              dateError: false,
              PHError: false,
              PHWeek: item.PHWeek ? item.PHWeek : null,
            });
          });
          adpGetTemplateData(adpItem, lessons, adpAllitems, items.length);
          // groups(adpAllitems);
          // adpGetAllOptions(adpAllitems);

          // setAdpActivityResponseData([...adpAllitems]);
          // setAdpData([...adpAllitems]);
          // setAdpMasterData([...adpAllitems]);
          // setAdpLoader("noLoader");
        } else {
          setNewDataFlag(true);
          sharepointWeb.lists
            .getByTitle(templateListName)
            .items.filter(`Types eq '${adpItem.Types}'`)
            .orderBy("ID", true)
            .top(5000)
            .get()
            .then((items) => {
              let count = 0;
              lessons.forEach((ls) => {
                items.forEach((item, index) => {
                  let PHErrorFlag =
                    item.MinHours && item.MaxHours
                      ? adpPHValidationFunction(
                          parseFloat(item.Hours ? item.Hours : 0),
                          item.MinHours,
                          item.MaxHours
                        )
                      : false;
                  // let datediff =
                  //   new Date(ls.EndDate).getDate() -
                  //   new Date(ls.StartDate).getDate() +
                  //   1;
                  // let Hours = item.Week ? datediff / 7 : item.Hours;
                  const date1: any = new Date(ls.StartDate);
                  const date2: any = new Date(ls.EndDate);
                  const diffTime = Math.abs(date2 - date1);
                  const diffDays =
                    Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
                  const Hours = item.Week ? diffDays / 7 : item.Hours;

                  adpAllitems.push({
                    OrderId: count++,
                    ID: count,
                    Steps: item.Title ? item.Title : "",
                    PH: Hours ? Hours : "",
                    MinPH: item.MinHours ? item.MinHours : "",
                    MaxPH: item.MaxHours ? item.MaxHours : "",
                    Project: adpItem.Project ? adpItem.Project : "",
                    LessonID: ls.ID,
                    Lesson: ls.Name ? ls.Name : "",
                    Start: moment(ls.StartDate).format(DateListFormat),
                    End: moment(ls.EndDate).format(DateListFormat),
                    Developer: {
                      name: ls.DeveloperName,
                      id: ls.DeveloperId,
                      email: ls.DeveloperEmail,
                    },
                    Status: "Scheduled",
                    IsCompleteStatus: false,
                    IsCompleteNew: false,
                    AH: 0,
                    dateError: false,
                    PHError: PHErrorFlag,
                    PHWeek: item.Week ? item.Week : null,
                  });
                });
              });
              groups(adpAllitems);
              adpGetAllOptions(adpAllitems);

              // setAdpActivityResponseData([...adpAllitems]);
              setAdpData([...adpAllitems]);
              setAdpMasterData([...adpAllitems]);
              setAdpLoader("noLoader");
            })
            .catch((err) => {
              adpErrorFunction(err, "adpGetData-getTemplateData");
            });
        }
      })
      .catch((err) => {
        adpErrorFunction(err, "adpGetData-getADPData");
      });
  };

  //!Update template in the database
  const adpGetTemplateData = (
    adpItem: any,
    lessons,
    adplistItems: any[],
    countLists
  ) => {
    let adpAllitems = adplistItems;
    sharepointWeb.lists
      .getByTitle(templateListName)
      .items.filter(`Types eq '${adpItem.Types}'`)
      .orderBy("ID", true)
      .top(5000)
      .get()
      .then((items) => {
console.log(items , "items in Activity Delivery Plan Template");


        let count = countLists;
        lessons.forEach((ls) => {
          let curLessondata = adplistItems.filter((arr) => {
            return arr.LessonID == ls.ID;
          });

          curLessondata.length == 0 &&
            items.forEach((item, index) => {
              let PHErrorFlag =
                item.MinHours && item.MaxHours
                  ? adpPHValidationFunction(
                      parseFloat(item.Hours ? item.Hours : 0),
                      item.MinHours,
                      item.MaxHours
                    )
                  : false;
              const date1: any = new Date(ls.StartDate);
              const date2: any = new Date(ls.EndDate);
              const diffTime = Math.abs(date2 - date1);
              const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
              const Hours = item.Week ? diffDays / 7 : item.Hours;

              adpAllitems.push({
                OrderId: count++,
                ID: 0,
                Steps: item.Title ? item.Title : "",
                PH: Hours ? Hours : "",
                MinPH: item.MinHours ? item.MinHours : "",
                MaxPH: item.MaxHours ? item.MaxHours : "",
                Project: adpItem.Project ? adpItem.Project : "",
                LessonID: ls.ID,
                Lesson: ls.Name ? ls.Name : "",
                Start: moment(ls.StartDate).format(DateListFormat),
                End: moment(ls.EndDate).format(DateListFormat),
                Developer: {
                  name: ls.DeveloperName,
                  id: ls.DeveloperId,
                  email: ls.DeveloperEmail,
                },
                Status: "Scheduled",
                IsCompleteStatus: false,
                IsCompleteNew: false,
                AH: 0,
                dateError: false,
                PHError: PHErrorFlag,
                PHWeek: item.Week ? item.Week : null,
              });
            });
        });
        groups(adpAllitems);
        adpGetAllOptions(adpAllitems);

        // setAdpActivityResponseData([...adpAllitems]);
        setAdpData([...adpAllitems]);
        setAdpMasterData([...adpAllitems]);
        setAdpLoader("noLoader");
      })
      .catch((err) => {
        adpErrorFunction(err, "adpGetData-getTemplateData");
      });
  };

  const adpGetAllOptions = (allItems: any) => {
    allItems.forEach((item: any) => {
      if (
        adpDrpDwnOptns.developerOptns.findIndex((developerOptn) => {
          return developerOptn.key == item.Developer.name;
        }) == -1 &&
        item.Developer.name
      ) {
        adpDrpDwnOptns.developerOptns.push({
          key: item.Developer.name,
          text: item.Developer.name,
        });
      }

      if (
        adpDrpDwnOptns.stepsOptns.findIndex((stepsOptn) => {
          return stepsOptn.key == item.Steps;
        }) == -1 &&
        item.Steps
      ) {
        adpDrpDwnOptns.stepsOptns.push({
          key: item.Steps,
          text: item.Steps,
        });
      }
      if (
        adpDrpDwnOptns.statusOptns.findIndex((statsOptn) => {
          return statsOptn.key == item.Status;
        }) == -1 &&
        item.Status
      ) {
        adpDrpDwnOptns.statusOptns.push({
          key: item.Status,
          text: item.Status,
        });
      }

      if (
        adpDrpDwnOptns.lessonOptns.findIndex((lessonOptn) => {
          return lessonOptn.key == item.Lesson;
        }) == -1 &&
        item.Lesson
      ) {
        adpDrpDwnOptns.lessonOptns.push({
          key: item.Lesson,
          text: item.Lesson,
        });
      }
    });

    let maxWeek =
      parseInt(adpFilters.year) == moment().year() ? moment().isoWeek() : 53;

    for (var i = 1; i <= maxWeek; i++) {
      adpDrpDwnOptns.weekOptns.push({
        key: i.toString(),
        text: i.toString(),
      });
    }
    for (var i = 2020; i <= moment().year(); i++) {
      adpDrpDwnOptns.yearOptns.push({
        key: i.toString(),
        text: i.toString(),
      });
    }

    let unsortedFilterKeys = adpSortingFilterKeys(adpDrpDwnOptns);
    setAdpDropDownOptions({ ...unsortedFilterKeys });
  };
  const adpSortingFilterKeys = (unsortedFilterKeys: any) => {
    const sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    if (
      unsortedFilterKeys.developerOptns.some((managerOptn) => {
        return (
          managerOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      unsortedFilterKeys.developerOptns.shift();
      let loginUserIndex = unsortedFilterKeys.developerOptns.findIndex(
        (user) => {
          return (
            user.text.toLowerCase() ==
            props.spcontext.pageContext.user.displayName.toLowerCase()
          );
        }
      );
      let loginUserData = unsortedFilterKeys.developerOptns.splice(
        loginUserIndex,
        1
      );

      unsortedFilterKeys.developerOptns.sort(sortFilterKeys);
      unsortedFilterKeys.developerOptns.unshift(loginUserData[0]);
      unsortedFilterKeys.developerOptns.unshift({ key: "All", text: "All" });
    } else {
      unsortedFilterKeys.developerOptns.shift();
      unsortedFilterKeys.developerOptns.sort(sortFilterKeys);
      unsortedFilterKeys.developerOptns.unshift({ key: "All", text: "All" });
    }

    unsortedFilterKeys.statusOptns.shift();
    unsortedFilterKeys.statusOptns.sort(sortFilterKeys);
    unsortedFilterKeys.statusOptns.unshift({ key: "All", text: "All" });

    unsortedFilterKeys.stepsOptns.shift();
    unsortedFilterKeys.stepsOptns.sort(sortFilterKeys);
    unsortedFilterKeys.stepsOptns.unshift({ key: "All", text: "All" });

    unsortedFilterKeys.lessonOptns.shift();
    unsortedFilterKeys.lessonOptns.sort(sortFilterKeys);
    unsortedFilterKeys.lessonOptns.unshift({ key: "All", text: "All" });

    return unsortedFilterKeys;
  };
  const adpListFilter = (key: string, option: any) => {
    let arrBeforeFilter = [...adpData];

    let tempFilterKeys = { ...adpFilters };
    tempFilterKeys[key] = option;

    if (tempFilterKeys.developer != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Developer.name == tempFilterKeys.developer;
      });
    }

    if (tempFilterKeys.step != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Steps == tempFilterKeys.step;
      });
    }
    if (tempFilterKeys.status != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Status == tempFilterKeys.status;
      });
    }
    if (tempFilterKeys.lesson) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Lesson.toLowerCase().includes(
          tempFilterKeys.lesson.toLowerCase()
        );
      });
    }

    if (tempFilterKeys.week != "All") {
      let year =
        tempFilterKeys.year == "All" ? moment().year() : tempFilterKeys.year;

      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        let start = moment(arr.Start, DateListFormat)
          .year()
          .toString()
          .concat(
            (
              "0" + moment(arr.Start, DateListFormat).isoWeek().toString()
            ).slice(-2)
          );
        let end = moment(arr.End, DateListFormat)
          .year()
          .toString()
          .concat(
            ("0" + moment(arr.End, DateListFormat).isoWeek().toString()).slice(
              -2
            )
          );
        let today = year
          .toString()
          .concat(("0" + tempFilterKeys.week.toString()).slice(-2));

        return (
          parseInt(today) >= parseInt(start) && parseInt(today) <= parseInt(end)
        );
      });
    }

    if (tempFilterKeys.year != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        let start = moment(arr.Start, DateListFormat).year().toString();

        let end = moment(arr.End, DateListFormat).year().toString();

        let today = tempFilterKeys.year.toString();

        return (
          parseInt(today) >= parseInt(start) && parseInt(today) <= parseInt(end)
        );
      });
    }

    groups([...arrBeforeFilter]);
    // setAdpActivityResponseData([...arrBeforeFilter]);
    setAdpFilters({ ...tempFilterKeys });
  };
  const overallPlannedHours = () => {
    let ph = 0;
    if (adpData.length > 0) {
      adpData.forEach((data) => {
        ph += data.PH ? data.PH : 0;
      });
    }
    return ph;
  };
  const overallActualHours = () => {
    let ah = 0;
    if (adpData.length > 0) {
      adpData.forEach((data) => {
        ah += data.AH ? data.AH : 0;
      });
    }
    return ah;
  };

  const adpActivityResponseHandler = (id: number, key: string, value: any) => {
    let tempDeveloper = [];

    let Index = adpData.findIndex((data) => data.OrderId == id);
    let disIndex = adpActivityResponseData.findIndex(
      (data) => data.OrderId == id
    );

    let adpBeforeData = adpData[Index];

    if (key == "Developer") {
      if (value) {
        tempDeveloper = allPeoples.filter((people) => {
          return people.ID == value;
        });
      }
    }

    let dateErrorFlag = adpDateValidationFunction(
      key == "Start"
        ? moment(value).format("YYYY/MM/DD")
        : moment(adpBeforeData.Start, DateListFormat).format("YYYY/MM/DD"),
      key == "End"
        ? moment(value).format("YYYY/MM/DD")
        : moment(adpBeforeData.End, DateListFormat).format("YYYY/MM/DD")
    );
    let PHErrorFlag =
      key == "PH"
        ? adpPHValidationFunction(
            parseFloat(value),
            adpBeforeData.MinPH,
            adpBeforeData.MaxPH
          )
        : adpBeforeData.PHError;

    const date1: any = new Date(key == "Start" ? value : adpBeforeData.Start);
    const date2: any = new Date(key == "End" ? value : adpBeforeData.End);
    const diffTime = Math.abs(date2 - date1);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    const Hours = adpBeforeData.PHWeek ? diffDays / 7 : adpBeforeData.PH;

    let adpOnchangeData = {
      LessonID: adpBeforeData.LessonID,
      OrderId: adpBeforeData.OrderId,
      ID: adpBeforeData.ID,
      Steps: adpBeforeData.Steps,
      PH: key == "PH" ? value : Hours,
      MinPH: adpBeforeData.MinPH,
      MaxPH: adpBeforeData.MaxPH,
      Project: adpBeforeData.Project,
      Lesson: adpBeforeData.Lesson,
      Start:
        key == "Start"
          ? moment(value).format(DateListFormat)
          : adpBeforeData.Start,
      End:
        key == "End" ? moment(value).format(DateListFormat) : adpBeforeData.End,
      Developer:
        key == "Developer"
          ? {
              name: tempDeveloper.length > 0 ? tempDeveloper[0].text : "",
              id: tempDeveloper.length > 0 ? tempDeveloper[0].ID : null,
              email:
                tempDeveloper.length > 0 ? tempDeveloper[0].secondaryText : "",
            }
          : adpBeforeData.Developer,
      Status: adpBeforeData.Status,
      IsCompleteStatus:
        key == "IsCompleteStatus" ? value : adpBeforeData.IsCompleteStatus,
      IsCompleteNew:
        key == "IsCompleteStatus" ? value : adpBeforeData.IsCompleteNew,
      dateError: dateErrorFlag,
      PHError: PHErrorFlag,
      PHWeek: adpBeforeData.PHWeek,
    };

    adpData[Index] = adpOnchangeData;
    adpActivityResponseData[disIndex] = adpOnchangeData;

    let isCompleteDetails = adpData.filter((arr) => {
      return arr.IsCompleteStatus == true;
    });
    adpData.length == isCompleteDetails.length
      ? setAdpIsCompleted(true)
      : setAdpIsCompleted(false);

    setAdpData([...adpData]);
    groups([...adpActivityResponseData]);
    // setAdpActivityResponseData([...adpActivityResponseData]);
  };
  const adpAddItem = () => {
    let successCount = 0;

    let completionDetails = adpData.filter((arr) => {
      return arr.IsCompleteStatus == true;
    });

    let completionValue =
      adpData.length > 0 && completionDetails.length > 0
        ? ((completionDetails.length / adpData.length) * 100).toFixed(2)
        : 0;

    adpData.forEach(async (response: any, index: number) => {
      let strDSNA: string = `${
        response.Developer.id ? response.Developer.id : null
      }-0`;

      let statusValue = response.IsCompleteStatus
        ? "Completed"
        : response.Status
        ? response.Status
        : null;

      let responseData = {
        ActivityPlanID: activityPlan_ID ? activityPlan_ID.toString() : "",
        Title: response.Steps ? response.Steps : "",
        PlannedHours: response.PH ? response.PH : 0,
        MinPH: response.MinPH ? response.MinPH : 0,
        MaxPH: response.MaxPH ? response.MaxPH : 0,
        ProjectVersion: activtyPlanItem[0].ProjectVersion
          ? activtyPlanItem[0].ProjectVersion
          : "V1",
        ProductVersion: activtyPlanItem[0].ProductVersion
          ? activtyPlanItem[0].ProductVersion
          : "V1",
        Product: activtyPlanItem[0].Product ? activtyPlanItem[0].Product : "",
        Types: activtyPlanItem[0].Types ? activtyPlanItem[0].Types : "",
        Project: activtyPlanItem[0].Project ? activtyPlanItem[0].Project : "",
        Lesson: response.Lesson ? response.Lesson : "",
        StartDate: response.Start
          ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
          : moment().format("YYYY-MM-DD"),
        EndDate: response.End
          ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
          : moment().format("YYYY-MM-DD"),
        DeveloperId: response.Developer.id ? response.Developer.id : null,
        // Status: "Scheduled",
        Status: statusValue,
        ActualHours: 0,
        OrderId: response.OrderId,
        LessonID: response.LessonID ? response.LessonID : null,
        PHWeek: response.PHWeek ? response.PHWeek : null,
        SPFxFilter: strDSNA,
      };

      // debugger;

      await sharepointWeb.lists
        .getByTitle(adpListName)
        .items.add(responseData)
        .then((item) => {
          successCount++;
          adpData[index].ID = item.data.Id;
          adpData[index].Status = statusValue;
          adpData[index].IsCompleteNew = false;

          if (adpData.length == successCount) {
            let apCompletedValue = adpData.filter((arr) => {
              return arr.IsCompleteStatus == true;
            });
            if (
              adpData.length == apCompletedValue.length &&
              activtyPlanItem[0].Status != "Completed"
            ) {
              sharepointWeb.lists
                .getByTitle(activityPlanListName)
                .items.getById(activityPlan_ID)
                .update({
                  Status: "Completed",
                  Completion: 100,
                  CompletedDate: moment().format("YYYY-MM-DD"),
                })
                .then((e) => {})
                .catch((err) => {
                  adpErrorFunction(err, "saveDPData-getAPItem");
                });
            } else {
              sharepointWeb.lists
                .getByTitle(activityPlanListName)
                .items.getById(activityPlan_ID)
                .update({
                  Completion: completionValue,
                })
                .then((e) => {})
                .catch((err) => {
                  adpErrorFunction(err, "saveDPData-getAPItem");
                });
            }

            const newData = _copyAndSort(adpData, "OrderId", false);

            adpGetAllOptions([...newData]);
            setAdpMasterData([...newData]);
            setAdpData([...newData]);
            groups([...newData]);

            // setAdpActivityResponseData([...adpData]);
            setNewDataFlag(false);
            setAdpItemAddFlag(true);
            setAdpEditFlag(false);
            setAdpLoader("noLoader");
            AddSuccessPopup();
          }
        })
        .catch((err) => {
          adpErrorFunction(err, "adpAddItem");
        });
    });
  };

  const adpUpdateItem_Old = () => {
    let responseDataArr = [];
    let newArr = [...adpData];
    let successCount = 0;

    let selected = [];

    adpData.forEach((response: any, index: number) => {
      let targetStatus = newArr.filter((arr) => {
        return arr.ID == response.ID;
      });

      let strDSNA: string = `${response.Developer.id}-${
        targetStatus[0].Status == "Completed" ? 1 : 0
      }`;

      let responseData = {
        ProjectVersion: activtyPlanItem[0].ProjectVersion
          ? activtyPlanItem[0].ProjectVersion
          : "V1",
        ProductVersion: activtyPlanItem[0].ProductVersion
          ? activtyPlanItem[0].ProductVersion
          : "V1",
        Product: activtyPlanItem[0].Product ? activtyPlanItem[0].Product : "",
        Project: activtyPlanItem[0].Project ? activtyPlanItem[0].Project : "",
        PlannedHours: response.PH ? response.PH : 0,
        StartDate: response.Start
          ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
          : null,
        EndDate: response.End
          ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
          : null,
        DeveloperId: response.Developer.id ? response.Developer.id : null,
        SPFxFilter: strDSNA,
      };

      responseDataArr.push(responseData);

      sharepointWeb.lists
        .getByTitle(adpListName)
        .items.getById(response.ID)
        .update(responseData)
        .then(() => {
          successCount++;
          let newDeveloperDetails = {};

          let targetIndex = newArr.findIndex((arr) => arr.ID == response.ID);
          let targetItem = newArr.filter((arr) => {
            return arr.ID == response.ID;
          });

          if (response.Developer.id) {
            let newDeveloper = allPeoples.filter((people) => {
              return people.ID == response.Developer.id;
            });
            newDeveloperDetails = {
              name: newDeveloper[0].text,
              id: newDeveloper[0].ID,
              email: newDeveloper[0].secondaryText,
            };
          } else {
            newDeveloperDetails = {
              name: null,
              id: null,
              email: null,
            };
          }

          newArr[targetIndex] = {
            OrderId: response.OrderId,
            ID: targetItem[0].ID ? targetItem[0].ID : "",
            Steps: targetItem[0].Steps ? targetItem[0].Steps : "",
            PH: response.PH ? response.PH : "",
            MinPH: targetItem[0].MinPH ? targetItem[0].MinPH : "",
            MaxPH: targetItem[0].MaxPH ? targetItem[0].MaxPH : "",
            Project: targetItem[0].Project ? targetItem[0].Project : "",
            LessonID: targetItem[0].LessonID ? targetItem[0].LessonID : null,
            Lesson: targetItem[0].Lesson ? targetItem[0].Lesson : "",
            Start: response.Start ? response.Start : targetItem[0].Start,
            End: response.End ? response.End : targetItem[0].End,
            Developer: newDeveloperDetails,
            Status: targetItem[0].Status ? targetItem[0].Status : "",
            AH: targetItem[0].AH ? targetItem[0].AH : "",
            dateError: false,
            PHError: false,
            PHWeek: targetItem[0].PHWeek ? targetItem[0].PHWeek : null,
          };

          let filteredPB = activityPB.filter((pb) => {
            return pb.ActivityDeliveryPlanID == newArr[targetIndex].ID;
          });

          selected.push([...filteredPB]);

          if (filteredPB.length > 0) {
            sharepointWeb.lists
              .getByTitle(activityPBListName)
              .items.getById(filteredPB[0].ID)
              .update({
                PlannedHours: response.PH ? response.PH : 0,
                StartDate: response.Start
                  ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
                  : null,
                EndDate: response.End
                  ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
                  : null,
                DeveloperId: response.Developer.id
                  ? response.Developer.id
                  : null,
              })
              .then((e) => {})
              .catch((err) => {
                adpErrorFunction(err, "adpUpdateItem-updateAPBList");
              });
          }

          if (adpActivityResponseData.length == successCount) {
            adpGetAllOptions(newArr);
            setAdpEditFlag(false);
            setAdpMasterData([...newArr]);
            setAdpLoader("noLoader");
            AddSuccessPopup();
          }
        })
        .catch((err) => {
          adpErrorFunction(err, "adpUpdateItem-updateATPList");
        });
    });
  };

  const adpUpdateItem = () => {
    let responseDataArr = [];
    let newArr = [...adpData];
    let successCount = 0;

    let completionDetails = adpData.filter((arr) => {
      return arr.IsCompleteStatus == true;
    });

    let completionValue =
      adpData.length > 0 && completionDetails.length > 0
        ? ((completionDetails.length / adpData.length) * 100).toFixed(2)
        : 0;

    let selected = [];

    adpData.forEach((response: any, index: number) => {
      if (response.ID != 0) {
        let targetStatus = newArr.filter((arr) => {
          return arr.ID == response.ID;
        });

        let strDSNA: string = `${response.Developer.id}-${
          targetStatus[0].Status == "Completed" ? 1 : 0
        }`;

        let statusValue = response.IsCompleteStatus
          ? "Completed"
          : response.Status
          ? response.Status
          : null;

        let responseData = {
          ProjectVersion: activtyPlanItem[0].ProjectVersion
            ? activtyPlanItem[0].ProjectVersion
            : "V1",
          ProductVersion: activtyPlanItem[0].ProductVersion
            ? activtyPlanItem[0].ProductVersion
            : "V1",
          Product: activtyPlanItem[0].Product ? activtyPlanItem[0].Product : "",
          Project: activtyPlanItem[0].Project ? activtyPlanItem[0].Project : "",
          PlannedHours: response.PH ? response.PH : 0,
          StartDate: response.Start
            ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
            : null,
          EndDate: response.End
            ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
            : null,
          DeveloperId: response.Developer.id ? response.Developer.id : null,
          SPFxFilter: strDSNA,
          OrderId: response.OrderId,
          Status: statusValue,
        };

        responseDataArr.push(responseData);

        sharepointWeb.lists
          .getByTitle(adpListName)
          .items.getById(response.ID)
          .update(responseData)
          .then(() => {
            successCount++;
            let newDeveloperDetails = {};

            let targetIndex = newArr.findIndex((arr) => arr.ID == response.ID);
            let targetItem = newArr.filter((arr) => {
              return arr.ID == response.ID;
            });

            if (response.Developer.id) {
              let newDeveloper = allPeoples.filter((people) => {
                return people.ID == response.Developer.id;
              });
              newDeveloperDetails = {
                name: newDeveloper[0].text,
                id: newDeveloper[0].ID,
                email: newDeveloper[0].secondaryText,
              };
            } else {
              newDeveloperDetails = {
                name: null,
                id: null,
                email: null,
              };
            }

            newArr[targetIndex] = {
              OrderId: index,
              ID: targetItem[0].ID ? targetItem[0].ID : "",
              Steps: targetItem[0].Steps ? targetItem[0].Steps : "",
              PH: response.PH ? response.PH : "",
              MinPH: targetItem[0].MinPH ? targetItem[0].MinPH : "",
              MaxPH: targetItem[0].MaxPH ? targetItem[0].MaxPH : "",
              Project: targetItem[0].Project ? targetItem[0].Project : "",
              LessonID: targetItem[0].LessonID ? targetItem[0].LessonID : null,
              Lesson: targetItem[0].Lesson ? targetItem[0].Lesson : "",
              Start: response.Start ? response.Start : targetItem[0].Start,
              End: response.End ? response.End : targetItem[0].End,
              Developer: newDeveloperDetails,
              // Status: targetItem[0].Status ? targetItem[0].Status : "",
              Status: statusValue,
              AH: targetItem[0].AH ? targetItem[0].AH : "",
              dateError: false,
              PHError: false,
              PHWeek: targetItem[0].PHWeek ? targetItem[0].PHWeek : null,
              IsCompleteStatus: response.IsCompleteStatus ? true : false,
              IsCompleteNew: false,
            };

            let filteredPB = activityPB.filter((pb) => {
              return pb.ActivityDeliveryPlanID == newArr[targetIndex].ID;
            });

            selected.push([...filteredPB]);

            if (filteredPB.length > 0) {
              sharepointWeb.lists
                .getByTitle(activityPBListName)
                .items.getById(filteredPB[0].ID)
                .update({
                  PlannedHours: response.PH ? response.PH : 0,
                  StartDate: response.Start
                    ? moment(response.Start, DateListFormat).format(
                        "YYYY-MM-DD"
                      )
                    : null,
                  EndDate: response.End
                    ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
                    : null,
                  DeveloperId: response.Developer.id
                    ? response.Developer.id
                    : null,
                })
                .then((e) => {})
                .catch((err) => {
                  adpErrorFunction(err, "adpUpdateItem-updateAPBList");
                });
            }

            if (adpData.length == successCount) {
              let apCompletedValue = adpData.filter((arr) => {
                return arr.IsCompleteStatus == true;
              });
              if (
                adpData.length == apCompletedValue.length &&
                activtyPlanItem[0].Status != "Completed"
              ) {
                sharepointWeb.lists
                  .getByTitle(activityPlanListName)
                  .items.getById(activityPlan_ID)
                  .update({
                    Status: "Completed",
                    Completion: 100,
                    CompletedDate: moment().format("YYYY-MM-DD"),
                  })
                  .then((e) => {})
                  .catch((err) => {
                    adpErrorFunction(err, "saveDPData-getAPItem");
                  });
              } else {
                sharepointWeb.lists
                  .getByTitle(activityPlanListName)
                  .items.getById(activityPlan_ID)
                  .update({
                    Completion: completionValue,
                  })
                  .then((e) => {})
                  .catch((err) => {
                    adpErrorFunction(err, "saveDPData-getAPItem");
                  });
              }

              const newData = _copyAndSort(newArr, "OrderId", false);

              adpGetAllOptions([...newData]);
              setAdpMasterData([...newData]);
              setAdpData([...newData]);
              groups([...newData]);

              // adpGetAllOptions(newArr);
              setAdpEditFlag(false);
              // setAdpMasterData([...newArr]);
              setAdpLoader("noLoader");
              AddSuccessPopup();
            }
          })
          .catch((err) => {
            adpErrorFunction(err, "adpUpdateItem-updateATPList");
          });
      } else {
        let strDSNA: string = `${
          response.Developer.id ? response.Developer.id : null
        }-0`;

        let responseData = {
          ActivityPlanID: activityPlan_ID ? activityPlan_ID.toString() : "",
          Title: response.Steps ? response.Steps : "",
          PlannedHours: response.PH ? response.PH : 0,
          MinPH: response.MinPH ? response.MinPH : 0,
          MaxPH: response.MaxPH ? response.MaxPH : 0,
          ProjectVersion: activtyPlanItem[0].ProjectVersion
            ? activtyPlanItem[0].ProjectVersion
            : "V1",
          ProductVersion: activtyPlanItem[0].ProductVersion
            ? activtyPlanItem[0].ProductVersion
            : "V1",
          Product: activtyPlanItem[0].Product ? activtyPlanItem[0].Product : "",
          Types: activtyPlanItem[0].Types ? activtyPlanItem[0].Types : "",
          Project: activtyPlanItem[0].Project ? activtyPlanItem[0].Project : "",
          Lesson: response.Lesson ? response.Lesson : "",
          StartDate: response.Start
            ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
            : moment().format("YYYY-MM-DD"),
          EndDate: response.End
            ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
            : moment().format("YYYY-MM-DD"),
          DeveloperId: response.Developer.id ? response.Developer.id : null,
          Status: "Scheduled",
          ActualHours: 0,
          OrderId: index,
          LessonID: response.LessonID ? response.LessonID : null,
          PHWeek: response.PHWeek ? response.PHWeek : null,
          SPFxFilter: strDSNA,
        };

        // debugger;

        sharepointWeb.lists
          .getByTitle(adpListName)
          .items.add(responseData)
          .then((item) => {
            successCount++;
            adpData[index].ID = item.data.Id;

            if (adpData.length == successCount) {
              let apCompletedValue = adpData.filter((arr) => {
                return arr.IsCompleteStatus == true;
              });
              if (
                adpData.length == apCompletedValue.length &&
                activtyPlanItem[0].Status != "Completed"
              ) {
                sharepointWeb.lists
                  .getByTitle(activityPlanListName)
                  .items.getById(activityPlan_ID)
                  .update({
                    Status: "Completed",
                    Completion: 100,
                    CompletedDate: moment().format("YYYY-MM-DD"),
                  })
                  .then((e) => {})
                  .catch((err) => {
                    adpErrorFunction(err, "saveDPData-getAPItem");
                  });
              } else {
                sharepointWeb.lists
                  .getByTitle(activityPlanListName)
                  .items.getById(activityPlan_ID)
                  .update({
                    Completion: completionValue,
                  })
                  .then((e) => {})
                  .catch((err) => {
                    adpErrorFunction(err, "saveDPData-getAPItem");
                  });
              }

              const newData = _copyAndSort(adpData, "OrderId", false);

              adpGetAllOptions([...newData]);
              setAdpMasterData([...newData]);
              setAdpData([...newData]);
              groups([...newData]);
              // setAdpActivityResponseData([...adpData]);
              setNewDataFlag(false);
              setAdpItemAddFlag(true);
              setAdpEditFlag(false);
              setAdpLoader("noLoader");
              AddSuccessPopup();
            }
          })
          .catch((err) => {
            adpErrorFunction(err, "adpAddItem");
          });
      }
    });
  };
  const adpDateValidationFunction = (startDate: any, EndDate: any) => {
    if (startDate != null && EndDate != null) {
      if (startDate > EndDate) {
        return true;
      } else {
        return false;
      }
    } else {
      return false;
    }
  };
  const adpPHValidationFunction = (val, min, max) => {
    if (val >= min && val <= max) {
      return false;
    } else {
      return true;
    }
  };
  const dateFormater = (date: Date): string => {
    return date ? moment(date).format("DD/MM/YYYY") : "";
  };
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
  const adpErrorFunction = (error: any, functionName: string) => {
    console.log(error);

    let response = {
      ComponentName: "Activity delivery plan",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setAdpLoader("noLoader");
        setAdpEditFlag(false);
        ErrorPopup();
        setAdpReRender(!adpReRender);
      }
    );
  };
  const AddSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Activity planner is successfully submitted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  const sortingFunction = (columnName, sortType): void => {
    let tempArr = adpData;
    let tempDisArr = adpActivityResponseData;

    const newDisData = _copyAndSort(
      tempDisArr,
      columnName,
      sortType == "desc" ? true : false
    );
    const newData = _copyAndSort(
      tempArr,
      columnName,
      sortType == "desc" ? true : false
    );

    setAdpData([...newData]);
    groups([...newDisData]);
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

  const groups = (records) => {
    let reOrderedRecords = [];

    let Uniquelessons = records.reduce(function (item, e1) {
      var matches = item.filter(function (e2) {
        return e1.Lesson === e2.Lesson;
      });

      if (matches.length == 0) {
        item.push(e1);
      }
      return item;
    }, []);

    Uniquelessons.forEach((ul) => {
      let curLesson = records.filter((arr) => {
        return arr.Lesson == ul.Lesson;
      });
      reOrderedRecords = reOrderedRecords.concat(curLesson);
    });
    groupsforDL(reOrderedRecords);
  };

  const groupsforDL = (records) => {
    let newRecords = [];
    records.forEach((rd, index) => {
      newRecords.push({
        Lesson: rd.Lesson,
        indexValue: index,
      });
    });

    let varGroup = [];
    let Uniquelessons = newRecords.reduce(function (item, e1) {
      var matches = item.filter(function (e2) {
        return e1.Lesson === e2.Lesson;
      });

      if (matches.length == 0) {
        item.push(e1);
      }
      return item;
    }, []);

    Uniquelessons.forEach((ul) => {
      let lessonLength = newRecords.filter((arr) => {
        return arr.Lesson == ul.Lesson;
      }).length;
      varGroup.push({
        key: ul.Lesson,
        name: ul.Lesson,
        startIndex: ul.indexValue,
        count: lessonLength,
      });
    });
    setAdpActivityResponseData([...records]);
    setgroup([...varGroup]);
  };
  //Function-Section Ends
  useEffect(() => {
    if (
      adpAutoSave &&
      adpEditFlag &&
      adpData.some((data) => data.dateError == true) == false
    ) {
      setTimeout(() => {
        newDataFlag
          ? document.getElementById("adpbtnSave").click()
          : document.getElementById("adpbtnUpdate").click();
      }, 300000);
    }
  }, [adpAutoSave]);

  useEffect(() => {
    setAdpLoader("startUpLoader");
    getActivityPlanItem();
    getActivityPBData();
    adpGetCurrentUserDetails();
  }, [adpReRender]);
  return (
    <>
      <div style={{ padding: "5px 15px" }}>
        {adpLoader == "startUpLoader" ? <CustomLoader /> : null}
        {/* Header-Section Starts */}
        <div className={styles.adpHeaderSection} style={{ paddingBottom: "0" }}>
          {/* Popup-Section Starts */}
          <div></div>
          {/* Popup-Section Ends */}
          <div className={styles.adpHeader} style={{ marginBottom: "15px" }}>
            <div className={styles.dpTitle}>
              <Icon
                iconName="NavigateBack"
                className={adpIconStyleClass.navArrow}
                onClick={() => {
                  adpAutoSave
                    ? confirm(
                        "You have unsaved changes, are you sure you want to leave?"
                      )
                      ? props.handleclick("ActivityPlan", null, "adp")
                      : null
                    : props.handleclick("ActivityPlan", null, "adp");
                }}
              />
              <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
                Activity planner
              </Label>
            </div>
            {/* <div style={{ display: "flex" }}>
              <Persona
                size={PersonaSize.size32}
                presence={PersonaPresence.none}
                imageUrl={
                  "/_layouts/15/userphoto.aspx?size=S&username=" +
                  `${
                    activtyPlanItem.length > 0
                      ? activtyPlanItem[0]["DeveloperDetails"].email
                      : ""
                  }`
                }
              />
              <Label>
                {activtyPlanItem.length > 0
                  ? activtyPlanItem[0]["DeveloperDetails"].name
                  : ""}
              </Label>
            </div> */}
          </div>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              flexWrap: "wrap",
            }}
          >
            <div
              className={styles.adpHeaderDetails}
              style={{ marginLeft: "-10px" }}
            >
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Project :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0
                    ? activtyPlanItem[0].Project +
                      " " +
                      activtyPlanItem[0].ProjectVersion
                    : ""}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Product :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0
                    ? activtyPlanItem[0].Product +
                      " " +
                      activtyPlanItem[0].ProductVersion
                    : ""}
                </Label>
              </div>
              {/* <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Status :</Label>
                <Label style={{ color: "#038387", marginRight: "-25px" }}>
                  {overallStatus()}
                </Label>
              </div> *
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Type :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0 ? activtyPlanItem[0].Types : ""}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Project :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0 ? activtyPlanItem[0].Project : ""}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>AH/PH :</Label>
                <Label style={{ color: "#038387" }}>
                  {overallActualHours()}/{overallPlannedHours()}
                </Label>
              </div> */}
            </div>
            <div style={{ display: "flex" }}>
              <div
                style={{
                  display: "flex",
                  justifyContent: "flex-end",
                  marginTop: 2,
                  marginRight: 20,
                }}
              >
                <div>
                  <Label style={{ marginRight: 5 }}>
                    Number of records :{" "}
                    <span style={{ color: "#038387" }}>
                      {adpActivityResponseData.length}
                    </span>
                  </Label>
                </div>
              </div>
              <Label
                onClick={() => {
                  generateExcel();
                }}
                style={{
                  backgroundColor: "#EBEBEB",
                  padding: "0 15px",
                  cursor: "pointer",
                  fontSize: "12px",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  borderRadius: "3px",
                  color: "#1D6F42",
                  height: 34,
                  marginRight: 10,
                }}
              >
                <Icon
                  style={{
                    color: "#1D6F42",
                    marginRight: 5,
                  }}
                  iconName="ExcelDocument"
                  className={adpIconStyleClass.export}
                />
                Export as XLS
              </Label>
              {adpEditFlag ? (
                <PrimaryButton
                  className={adpbuttonStyleClass.buttonPrimary}
                  iconProps={cancelIcon}
                  onClick={() => {
                    setAdpEDSort("");
                    setAdpSDSort("");
                    setAdpEditFlag(false);
                    setAdpAutoSave(false);
                    setAdpData([...adpMasterData]);
                    groups([...adpMasterData]);
                    // setAdpActivityResponseData([...adpMasterData]);

                    setAdpFilters({ ...adpFilterKeys });
                  }}
                >
                  Cancel
                </PrimaryButton>
              ) : (
                <PrimaryButton
                  className={adpbuttonStyleClass.buttonPrimary}
                  iconProps={editIcon}
                  onClick={() => {
                    setAdpEditFlag(true);
                    setAdpAutoSave(true);
                  }}
                >
                  Edit
                </PrimaryButton>
              )}
              {newDataFlag == true && adpItemAddFlag == false ? (
                <PrimaryButton
                  id="adpbtnSave"
                  iconProps={saveIcon}
                  className={
                    adpEditFlag &&
                    adpData.some(
                      (data) => data.dateError == true || data.PHError == true
                    ) == false
                      ? adpbuttonStyleClass.buttonSecondary
                      : styles.adpSaveBtnDisabled
                  }
                  disabled={
                    adpEditFlag &&
                    adpData.some(
                      (data) => data.dateError == true || data.PHError == true
                    ) == false
                      ? false
                      : true
                  }
                  onClick={() => {
                    if (adpEditFlag) {
                      setAdpAutoSave(false);

                      let isCompletedData = adpData.filter((arr) => {
                        return arr.IsCompleteNew == true;
                      });
                      if (isCompletedData.length > 0) {
                        setAdpConfirmationPopup({
                          condition: true,
                          isNew: true,
                        });
                      } else {
                        setAdpLoader("startUpLoader");
                        adpAddItem();
                      }
                      // setAdpLoader("startUpLoader");
                      // adpAddItem();
                    }
                  }}
                >
                  {adpLoader == "saveLoader" ? <Spinner /> : <>Save</>}
                </PrimaryButton>
              ) : (
                <PrimaryButton
                  id="adpbtnUpdate"
                  iconProps={saveIcon}
                  className={
                    adpEditFlag &&
                    adpData.some(
                      (data) => data.dateError == true || data.PHError == true
                    ) == false
                      ? adpbuttonStyleClass.buttonSecondary
                      : styles.adpSaveBtnDisabled
                  }
                  disabled={
                    adpEditFlag &&
                    adpData.some(
                      (data) => data.dateError == true || data.PHError == true
                    ) == false
                      ? false
                      : true
                  }
                  onClick={() => {
                    if (
                      !adpData.some(
                        (data) => data.dateError == true || data.PHError == true
                      )
                    ) {
                      if (adpEditFlag) {
                        setAdpAutoSave(false);

                        let isCompletedData = adpData.filter((arr) => {
                          return arr.IsCompleteNew == true;
                        });
                        if (isCompletedData.length > 0) {
                          setAdpConfirmationPopup({
                            condition: true,
                            isNew: false,
                          });
                        } else {
                          setAdpLoader("startUpLoader");
                          adpUpdateItem();
                        }

                        // setAdpLoader("startUpLoader");
                        // adpUpdateItem();
                      }
                    }
                  }}
                >
                  {adpLoader == "updateLoader" ? <Spinner /> : <>Save</>}
                </PrimaryButton>
              )}
              <Icon
                iconName="Link12"
                className={adpIconStyleClass.link}
                onClick={() => {
                  adpAutoSave
                    ? confirm(
                        "You have unsaved changes, are you sure you want to leave?"
                      )
                      ? props.handleclick(
                          "ActivityProductionBoard",
                          activityPlan_ID,
                          "ADP"
                        )
                      : null
                    : props.handleclick(
                        "ActivityProductionBoard",
                        activityPlan_ID,
                        "ADP"
                      );
                }}
              />
            </div>
          </div>
          {/* Header-Section Ends */}
          {/* Filter-Section Starts */}
          <div>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                marginTop: "-5px",
                marginBottom: "10px",
                flexWrap: "wrap",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "flex-start",
                  flexWrap: "wrap",
                }}
              >
                {/* <div>
                  <Label styles={adpLabelStyles}>Section</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.lesson != "All"
                        ? adpActiveDropdownStyles
                        : adpDropdownStyles
                    }
                    options={adpDropDownOptions.lessonOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      adpListFilter("lesson", option["key"]);
                    }}
                    selectedKey={adpFilters.lesson}
                  />
                </div> */}
                <div>
                  <Label styles={adpLabelStyles}>Section</Label>
                  <SearchBox
                    styles={
                      adpFilters.lesson
                        ? adpActiveSearchBoxStyles
                        : adpSearchBoxStyles
                    }
                    value={adpFilters.lesson}
                    onChange={(e, value) => {
                      adpListFilter("lesson", value);
                    }}
                  />
                </div>
                <div>
                  <Label styles={adpLabelStyles}>Steps</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.step != "All"
                        ? adpActiveDropdownStyles
                        : adpDropdownStyles
                    }
                    options={adpDropDownOptions.stepsOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      adpListFilter("step", option["key"]);
                    }}
                    selectedKey={adpFilters.step}
                  />
                </div>
                <div>
                  <Label styles={adpLabelStyles}>Status</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.status != "All"
                        ? adpActiveDropdownStyles
                        : adpDropdownStyles
                    }
                    options={adpDropDownOptions.statusOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      adpListFilter("status", option["key"]);
                    }}
                    selectedKey={adpFilters.status}
                  />
                </div>
                <div>
                  <Label styles={adpLabelStyles}>Developer</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.developer != "All"
                        ? adpActiveDropdownStyles
                        : adpDropdownStyles
                    }
                    options={adpDropDownOptions.developerOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      adpListFilter("developer", option["key"]);
                    }}
                    selectedKey={adpFilters.developer}
                  />
                </div>
                <div>
                  <Label
                    style={{
                      width: 75,
                    }}
                    styles={adpLabelStyles}
                  >
                    Year
                  </Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.year != "All"
                        ? adpActiveShortDropdownStyles
                        : adpShortDropdownStyles
                    }
                    options={adpDropDownOptions.yearOptns}
                    onChange={(e, option: any) => {
                      adpListFilter("year", option["key"]);
                    }}
                    selectedKey={adpFilters.year}
                  />
                </div>
                <div>
                  <Label
                    style={{
                      width: 75,
                    }}
                    styles={adpLabelStyles}
                  >
                    Week
                  </Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.week != "All"
                        ? adpActiveShortDropdownStyles
                        : adpShortDropdownStyles
                    }
                    options={adpDropDownOptions.weekOptns}
                    onChange={(e, option: any) => {
                      adpListFilter("week", option["key"]);
                    }}
                    selectedKey={adpFilters.week}
                  />
                </div>

                <div>
                  <Label
                    style={{
                      width: 65,
                    }}
                    styles={adpLabelStyles}
                  >
                    Start date
                  </Label>
                  <div
                    style={{
                      display: "flex",
                      marginTop: 5,
                      marginRight: 15,
                    }}
                  >
                    <button
                      style={{
                        backgroundColor: "#038387",
                        border: 0,
                        borderRadius: 10,
                        padding: "0.25rem  0.5rem",
                        marginRight: 10,
                        cursor: "pointer",
                      }}
                    >
                      <Icon
                        title={"asc"}
                        style={{
                          color: "#fff",
                          fontSize: adpSDSort == "asc" ? 20 : 16,
                          fontWeight: adpSDSort == "asc" ? "bold" : "normal",
                        }}
                        iconName="SortUp"
                        onClick={() => {
                          setAdpEDSort("");
                          setAdpSDSort("asc");
                          sortingFunction("Start", "asc");
                        }}
                      />
                      <Icon
                        title={"desc"}
                        style={{
                          color: "#fff",
                          fontSize: adpSDSort == "desc" ? 20 : 16,
                          fontWeight: adpSDSort == "desc" ? "bold" : "normal",
                        }}
                        iconName="SortDown"
                        onClick={() => {
                          setAdpEDSort("");
                          setAdpSDSort("desc");
                          sortingFunction("Start", "desc");
                        }}
                      />
                    </button>
                  </div>
                </div>
                <div>
                  <Label
                    style={{
                      width: 65,
                    }}
                    styles={adpLabelStyles}
                  >
                    End date
                  </Label>
                  <div
                    style={{
                      display: "flex",
                      marginTop: 5,
                      marginRight: 15,
                    }}
                  >
                    <button
                      style={{
                        backgroundColor: "#038387",
                        border: 0,
                        borderRadius: 10,
                        padding: "0.25rem  0.5rem",
                        marginRight: 10,
                        cursor: "pointer",
                      }}
                    >
                      <Icon
                        title={"asc"}
                        style={{
                          color: "#fff",
                          fontSize: adpEDSort == "asc" ? 20 : 16,
                          fontWeight: adpEDSort == "asc" ? "bold" : "normal",
                        }}
                        iconName="SortUp"
                        onClick={() => {
                          setAdpSDSort("");
                          setAdpEDSort("asc");
                          sortingFunction("End", "asc");
                        }}
                      />
                      <Icon
                        title={"desc"}
                        style={{
                          color: "#fff",
                          fontSize: adpEDSort == "desc" ? 20 : 16,
                          fontWeight: adpEDSort == "desc" ? "bold" : "normal",
                        }}
                        iconName="SortDown"
                        onClick={() => {
                          setAdpSDSort("");
                          setAdpEDSort("desc");
                          sortingFunction("End", "desc");
                        }}
                      />
                    </button>
                  </div>
                </div>
                <div>
                  <Label style={{ width: 60 }} styles={adpLabelStyles}>
                    Complete
                  </Label>
                  <Checkbox
                    styles={{
                      root: { marginTop: 3, width: 50 },
                    }}
                    disabled={!adpEditFlag ? true : false}
                    checked={AdpIsCompleted}
                    onChange={(ev) => {
                      setAdpIsCompleted(!AdpIsCompleted);
                      adpData.forEach((item, Index) => {
                        let dpBeforeData = adpData[Index];
                        let dpOnchangeData = [
                          {
                            OrderId: dpBeforeData.OrderId,
                            ID: dpBeforeData.ID,
                            Steps: dpBeforeData.Steps,
                            PH: dpBeforeData.PH,
                            MinPH: dpBeforeData.MinPH,
                            MaxPH: dpBeforeData.MaxPH,
                            Project: dpBeforeData.Project,
                            LessonID: dpBeforeData.LessonID,
                            IsCompleteStatus:
                              item.Status != "Completed"
                                ? ev.target["checked"]
                                : dpBeforeData.IsCompleteStatus,
                            Lesson: dpBeforeData.Lesson,
                            Start: dpBeforeData.Start,
                            End: dpBeforeData.End,
                            Developer: dpBeforeData.Developer,
                            Status: dpBeforeData.Status,
                            AH: dpBeforeData.AH,
                            dateError: dpBeforeData.dateError,
                            PHError: dpBeforeData.PHError,
                            PHWeek: dpBeforeData.PHWeek,
                            IsCompleteNew:
                              item.Status != "Completed"
                                ? ev.target["checked"]
                                : dpBeforeData.IsCompleteNew,
                          },
                        ];
                        adpData[Index] = dpOnchangeData[0];
                      });

                      adpActivityResponseData.forEach((item, Index) => {
                        let dpBeforeData = adpActivityResponseData[Index];
                        let dpOnchangeData = [
                          {
                            OrderId: dpBeforeData.OrderId,
                            ID: dpBeforeData.ID,
                            Steps: dpBeforeData.Steps,
                            PH: dpBeforeData.PH,
                            MinPH: dpBeforeData.MinPH,
                            MaxPH: dpBeforeData.MaxPH,
                            Project: dpBeforeData.Project,
                            LessonID: dpBeforeData.LessonID,
                            IsCompleteStatus:
                              item.Status != "Completed"
                                ? ev.target["checked"]
                                : dpBeforeData.IsCompleteStatus,
                            Lesson: dpBeforeData.Lesson,
                            Start: dpBeforeData.Start,
                            End: dpBeforeData.End,
                            Developer: dpBeforeData.Developer,
                            Status: dpBeforeData.Status,
                            AH: dpBeforeData.AH,
                            dateError: dpBeforeData.dateError,
                            PHError: dpBeforeData.PHError,
                            PHWeek: dpBeforeData.PHWeek,
                            IsCompleteNew:
                              item.Status != "Completed"
                                ? ev.target["checked"]
                                : dpBeforeData.IsCompleteNew,
                          },
                        ];
                        adpActivityResponseData[Index] = dpOnchangeData[0];
                      });

                      setAdpData([...adpData]);
                      setAdpActivityResponseData([...adpActivityResponseData]);
                    }}
                  />
                </div>
                <div>
                  <Icon
                    iconName="Refresh"
                    title="Click to reset"
                    className={adpIconStyleClass.refresh}
                    onClick={() => {
                      if (adpAutoSave) {
                        if (
                          confirm(
                            "You have unsaved changes, are you sure you want to leave?"
                          )
                        ) {
                          setAdpEDSort("");
                          setAdpSDSort("");
                          groups(adpMasterData);
                          // setAdpActivityResponseData(adpMasterData);
                          setAdpData([...adpMasterData]);
                          adpGetAllOptions(adpMasterData);
                          setAdpFilters({ ...adpFilterKeys });
                        }
                      } else {
                        setAdpEDSort("");
                        setAdpSDSort("");
                        groups(adpMasterData);
                        // setAdpActivityResponseData(adpMasterData);
                        setAdpData([...adpMasterData]);
                        adpGetAllOptions(adpMasterData);
                        setAdpFilters({ ...adpFilterKeys });
                      }
                    }}
                  />
                </div>
                <div>
                  <div
                    style={{
                      // display: "flex",
                      // justifyContent: "flex-end",
                      // marginLeft: "20px",
                      marginTop: "38px",
                    }}
                  >
                    {adpEditFlag &&
                    adpData.some((data) => data.dateError == true) ? (
                      <Label
                        style={{
                          marginRight: 5,
                        }}
                        className={adpCommonStyles.dateGridValidationErrorLabel}
                      >
                        *Given end date should not be earlier than the start
                        date
                      </Label>
                    ) : null}
                    {adpEditFlag &&
                    adpData.some((data) => data.PHError == true) ? (
                      <Label
                        style={{
                          marginRight: 5,
                        }}
                        className={adpCommonStyles.dateGridValidationErrorLabel}
                      >
                        *Please enter valid hours(PH)
                      </Label>
                    ) : null}
                  </div>
                </div>
              </div>
            </div>
          </div>
          {/* Filter-Section Ends */}
        </div>

        {/* Body-Section Starts */}

        <div>
          {/* dont remove */}
          {/* <input
            id="forFocus"
            type="text"
            style={{
              width: 0,
              height: 0,
              border: "none",
              position: "absolute",
              top: 0,
              left: 0,
              padding: 0,
            }}
          /> */}
        </div>
        <div
          className={styles.scrollTop}
          onClick={() => {
            document.querySelector("#forFocus")["focus"]();
          }}
        >
          <Icon iconName="Up" style={{ color: "#fff" }} />
        </div>
        <div>
          {/* DetailList-Section Starts */}
          <div>
            {
              <DetailsList
                items={adpActivityResponseData}
                columns={_adpColumns}
                groups={group}
                groupProps={{
                  showEmptyGroups: true,
                }}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                styles={
                  adpActivityResponseData.length > 0
                    ? gridStyles
                    : noDatagridStyles
                }
                data-is-scrollable={true}
                onShouldVirtualize={() => {
                  return false;
                }}
                onRenderRow={(data, defaultRender) => (
                  <div>
                    {defaultRender({
                      ...data,
                      styles: {
                        root: {
                          background:
                            data.item.IsNew == true
                              ? "linear-gradient(90deg, rgba(250,163,50,0.1491947120645133) 35%, rgba(3,131,135,0.14639359161633403) 100%)"
                              : "#fff",
                          selectors: {
                            "&:hover": {
                              background:
                                data.item.IsNew == true
                                  ? "linear-gradient(270deg, rgba(250,163,50,0.19961488013174022) 35%, rgba(3,131,135,0.19961488013174022) 100%)"
                                  : "#f3f2f1",
                            },
                          },
                        },
                      },
                    })}
                  </div>
                )}
              />
            }
          </div>
          {adpActivityResponseData.length == 0 ? (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                marginTop: "15px",
              }}
            >
              <Label style={{ color: "#2392B2" }}>No Data Found !!!</Label>
            </div>
          ) : null}
          {/* DetailList-Section Ends */}
        </div>

        <div>
          <Modal isOpen={AdpConfirmationPopup.condition} isBlocking={true}>
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
                <Label className={styles.deletePopupTitle}>Confirmation</Label>
                <Label
                  style={{
                    padding: "5px 20px",
                  }}
                  className={styles.deletePopupDesc}
                >
                  Are you sure want to mark as completed?
                </Label>
              </div>
            </div>
            <div className={styles.apDeletePopupBtnSection}>
              <button
                onClick={(_) => {
                  setAdpConfirmationPopup({ condition: false, isNew: false });
                  // saveDPData();
                  setAdpLoader("startUpLoader");
                  AdpConfirmationPopup.isNew ? adpAddItem() : adpUpdateItem();
                }}
                className={styles.apDeletePopupYesBtn}
              >
                Yes
              </button>
              <button
                onClick={(_) => {
                  setAdpConfirmationPopup({ condition: false, isNew: false });
                }}
                className={styles.apDeletePopupNoBtn}
              >
                No
              </button>
            </div>
          </Modal>
        </div>

        {/* Body-Section Ends */}
      </div>
    </>
  );
};

export default ActivityDeliveryPlan;
