import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from "@pnp/sp/webs";
import {
  IColumn,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  Modal,
  DatePicker,
  NormalPeoplePicker,
  PrimaryButton,
  ChoiceGroup,
  TextField,
  ITextFieldStyles,
  Checkbox,
  Spinner,
  TooltipHost,
  TooltipOverflowMode,
} from "@fluentui/react";

import Service from "../components/Services";

import { FontIcon, IIconProps } from "@fluentui/react/lib/Icon";
import "../ExternalRef/styleSheets/Styles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import { IDetailsListStyles } from "office-ui-fabric-react";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

const saveIcon: IIconProps = { iconName: "Save" };
const editIcon: IIconProps = { iconName: "Edit" };
const cancelIcon: IIconProps = { iconName: "Cancel" };
const MoveUp = require("../ExternalRef/assets/moveUp.png");
const MoveDown = require("../ExternalRef/assets/moveDown.png");

//Sorting
let sortDpDataArr = [];
let sortDpDisplayArr = [];
let sortDpUpdate = false;
let gblDeliveryPlanTemplate = [];
let DateListFormat = "DD/MM/YYYY";
let DatePickerFormat = "YYYY-MM-DDT14:00:00Z";

const DeliveryPlan = (props: any) => {
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;
  let Ap_AnnualPlanId = props.AnnualPlanId;
  let Dp_Year = moment().year();
  let Dp_WeekNumber = moment().isoWeek();
  let loggeduseremail = props.spcontext.pageContext.user.email;

  // Items in Detail List
  let _dpAllitems = [];
  const allPeoples = props.peopleList;

  const _dpColumns = [
    {
      key: "Column1",
      name: "Source",
      fieldName: "Source",
      minWidth: 20,
      maxWidth: 60,
      onColumnClick: (ev, column) => {
        !sortDpUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item, index) =>
        index == 0 ? (
          <div>
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
            {item.Source}
          </div>
        ) : (
          <div>{item.Source}</div>
        ),
    },
    {
      key: "Column2",
      name: "Activity",
      fieldName: "Title",
      minWidth: 50,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        !sortDpUpdate ? _onColumnClick(ev, column) : null;
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
      key: "Column4",
      name: "N/A",
      fieldName: "NotApplicable",
      minWidth: 35,
      maxWidth: 50,
      onRender: (item, Index) => (
        <Checkbox
          styles={{ root: { marginTop: 3 } }}
          data-id={item.ID}
          disabled={!dpUpdate ? true : false}
          checked={item.NotApplicable}
          onChange={(ev) => {
            dpOnchangeItems(item.RefId, "NotApplicable", ev.target["checked"]);
          }}
        />
      ),
    },
    {
      key: "Column5",
      name: "N/A(C)",
      fieldName: "NotApplicableManager",
      minWidth: 40,
      maxWidth: 50,
      onRender: (item, Index) => (
        <Checkbox
          styles={{ root: { marginTop: 3 } }}
          data-id={item.ID}
          disabled={
            dpUpdate &&
            (apCurrentData.length > 0 && loggeduseremail != ""
              ? apCurrentData[0].ProjectOwnerEmail == loggeduseremail
              : null)
              ? false
              : true
          }
          checked={item.NotApplicableManager}
          onChange={(ev) => {
            dpOnchangeItems(
              item.RefId,
              "NotApplicableManager",
              ev.target["checked"]
            );
          }}
        />
      ),
    },
    {
      key: "Column6",
      name: "IsComplete",
      fieldName: "IsCompleteStatus",
      minWidth: 85,
      maxWidth: 85,
      onRender: (item, Index) => (
        <Checkbox
          styles={{ root: { marginTop: 3 } }}
          data-id={item.ID}
          disabled={!dpUpdate || item.Status == "Completed" ? true : false}
          checked={item.IsCompleteStatus}
          onChange={(ev) => {
            dpOnchangeItems(
              item.RefId,
              "IsCompleteStatus",
              ev.target["checked"]
            );
          }}
        />
      ),
    },
    {
      key: "Column7",
      name: "Hours",
      fieldName: "PlannedHours",
      minWidth: 40,
      maxWidth: 60,
      onColumnClick: (ev, column) => {
        !sortDpUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item, Index) => (
        <TextField
          styles={{
            root: {
              selectors: {
                ".ms-TextField-fieldGroup": {
                  borderRadius: 4,
                  border: "1px solid",
                  height: 24,
                  input: {
                    borderRadius: 4,
                  },
                },
              },
            },
          }}
          data-id={item.ID}
          disabled={!dpUpdate || item.Source == "DP" ? true : false}
          value={item.PlannedHours}
          onChange={(e: any) => {
            dpOnchangeItems(item.RefId, "PlannedHours", e.target.value);
          }}
        />
      ),
    },
    {
      key: "Column8",
      name: "Start date",
      fieldName: "StartDate",
      minWidth: 100,
      maxWidth: 120,
      onColumnClick: (ev, column) => {
        !sortDpUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item, Index) => (
        <DatePicker
          data-id={item.ID}
          placeholder="Select a date..."
          ariaLabel="Select a date"
          formatDate={dateFormater}
          styles={{
            textField: {
              transform: "translateY(3px)",
              selectors: {
                ".ms-TextField-fieldGroup": {
                  borderColor: item.DateError ? "#d0342c" : "#000",
                  borderRadius: 4,
                  border: "1px solid",
                  height: 23,
                  input: {
                    borderRadius: 4,
                  },
                },
                ".ms-TextField-field": {
                  color: item.DateError ? "#d0342c" : "#000",
                },
                ".ms-DatePicker-event--without-label": {
                  color: item.DateError ? "#d0342c" : "#000",
                  paddingTop: 3,
                },
              },
            },
            readOnlyTextField: {
              lineHeight: 22,
            },
          }}
          value={
            item.StartDate
              ? new Date(
                  moment(item.StartDate, DateListFormat).format(
                    DatePickerFormat
                  )
                )
              : new Date()
          }
          disabled={!dpUpdate ? true : false}
          onSelectDate={(value: any) => {
            dpOnchangeItems(item.RefId, "StartDate", value);
            let refIndex = dpData.findIndex((obj) => obj.RefId == item.RefId);
            if (
              moment(value).format("YYYY/MM/DD") <=
              moment(item.EndDate, DateListFormat).format("YYYY/MM/DD")
            ) {
              dpData[refIndex].DateError = false;
              dpDateErrorFunction();
            } else {
              dpData[refIndex].DateError = true;
              dpDateErrorFunction();
            }
          }}
        />
      ),
    },
    {
      key: "Column9",
      name: "End date",
      fieldName: "EndDate",
      minWidth: 100,
      maxWidth: 120,
      onColumnClick: (ev, column) => {
        !sortDpUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item, Index) => (
        <DatePicker
          data-id={item.ID}
          placeholder="Select a date..."
          ariaLabel="Select a date"
          formatDate={dateFormater}
          value={
            item.EndDate
              ? new Date(
                  moment(item.EndDate, DateListFormat).format(DatePickerFormat)
                )
              : new Date()
          }
          disabled={!dpUpdate ? true : false}
          styles={{
            textField: {
              transform: "translateY(3px)",
              selectors: {
                ".ms-TextField-fieldGroup": {
                  borderColor: item.DateError ? "#d0342c" : "#000",
                  borderRadius: 4,
                  border: "1px solid",
                  height: 23,
                  input: {
                    borderRadius: 4,
                  },
                },
                ".ms-TextField-field": {
                  color: item.DateError ? "#d0342c" : "#000",
                },
                ".ms-DatePicker-event--without-label": {
                  color: item.DateError ? "#d0342c" : "#000",
                  paddingTop: 3,
                },
              },
            },
            readOnlyTextField: {
              lineHeight: 22,
            },
          }}
          onSelectDate={(value: any) => {
            dpOnchangeItems(item.RefId, "EndDate", value);
            let refIndex = dpData.findIndex((obj) => obj.RefId == item.RefId);
            if (
              moment(item.StartDate, DateListFormat).format("YYYY/MM/DD") <=
              moment(value).format("YYYY/MM/DD")
            ) {
              dpData[refIndex].DateError = false;
              dpDateErrorFunction();
            } else {
              dpData[refIndex].DateError = true;
              dpDateErrorFunction();
            }
          }}
        />
      ),
    },
    {
      key: "Column10",
      name: "Status",
      fieldName: "Status",
      minWidth: 70,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        !sortDpUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item) => (
        <div /*style={{ marginTop: "0.2rem" }}*/>
          {item.Status == "Completed" ? (
            <div className={dpstatusStyleClass.completed}>{item.Status}</div>
          ) : item.Status == "Scheduled" ? (
            <div className={dpstatusStyleClass.scheduled}>{item.Status}</div>
          ) : item.Status == "On schedule" ? (
            <div className={dpstatusStyleClass.onSchedule}>{item.Status}</div>
          ) : item.Status == "Behind schedule" ? (
            <div className={dpstatusStyleClass.behindScheduled}>
              {item.Status}
            </div>
          ) : item.Status == "On hold" ? (
            <div className={dpstatusStyleClass.Onhold}>{item.Status}</div>
          ) : (
            ""
          )}
        </div>
      ),
    },
    {
      key: "Column11",
      name: "Developer",
      fieldName: "Developer",
      minWidth: 200,
      maxWidth: 350,
      onColumnClick: (ev, column) => {
        !sortDpUpdate ? _onColumnClick(ev, column) : null;
      },
      onRender: (item, Index) => (
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
                },
              },
            },
          }}
          data-id={item.ID}
          onResolveSuggestions={GetUserDetails}
          itemLimit={1}
          disabled={!dpUpdate ? true : false}
          selectedItems={
            item.DeveloperId != null
              ? [
                  allPeoples.filter((people) => {
                    return people.ID == item.DeveloperId;
                  })[0],
                ]
              : []
          }
          onChange={(selectedUser) => {
            selectedUser.length != 0
              ? dpOnchangeItems(
                  item.RefId,
                  "DeveloperId",
                  selectedUser[0]["ID"]
                )
              : dpOnchangeItems(item.RefId, "DeveloperId", null);
          }}
        />
      ),
    },
    {
      key: "Column12",
      name: "TL",
      fieldName: "TLink",
      minWidth: 20,
      maxWidth: 30,
      onRender: (item) => {
        let templatelink;
        let curPhase =
          gblDeliveryPlanTemplate.length > 0
            ? gblDeliveryPlanTemplate.filter((arr) => {
                return arr.Title == item.Title;
              })
            : [];
        if (curPhase.length > 0) {
          templatelink = curPhase[0].TemplateUrl + "?web=1";
        } else {
          templatelink = "No data Found";
        }
        return (
          <>
            {item.Source == "DP" ? (
              <a data-interception="off" target="_blank" href={templatelink}>
                <Icon
                  title="Template link"
                  iconName="NavigateExternalInline"
                  className={dpiconStyleClass.link}
                  style={{ color: "#038387" }}
                />
              </a>
            ) : null}
          </>
        );
      },
    },
    {
      key: "Column13",
      name: "PBL",
      fieldName: "PBLink",
      minWidth: 30,
      maxWidth: 40,
      onRender: (item) => {
        let phaseID;
        let curPhase =
          gblDeliveryPlanTemplate.length > 0
            ? gblDeliveryPlanTemplate.filter((arr) => {
                return arr.Title == item.Title;
              })
            : [];
        if (curPhase.length > 0) {
          phaseID = curPhase[0].PhasesId;
        } else {
          phaseID = 0;
        }
        return (
          <>
            {item.Source == "DP" ? (
              <a
                data-interception="off"
                target="_blank"
                href={
                  props.playbookURL + Ap_AnnualPlanId + "&PhaseId=" + phaseID
                }
              >
                <Icon
                  title="Productionboard link"
                  iconName="NavigateExternalInline"
                  className={dpiconStyleClass.link}
                  style={{ color: "#038387" }}
                />
              </a>
            ) : null}
          </>
        );
      },
    },
    {
      key: "Column14",
      name: "Action",
      fieldName: "Arrow",
      minWidth: 100,
      maxWidth: 100,
      onRender: (item, index) => (
        <div
          style={{
            display: "flex",
            flexDirection: "row",
            alignItems: "center",
            marginTop: 6,
          }}
        >
          {index != 0 ? (
            <img
              title="Move up"
              style={{
                cursor:
                  dpUpdate && dpData.length == dpDisplayData.length
                    ? "pointer"
                    : "default",
              }}
              src={`${MoveUp}`}
              width={36}
              height={36}
              onClick={(_) => {
                dpUpdate && dpData.length == dpDisplayData.length
                  ? dpArrows(index, "Up")
                  : null;
              }}
            />
          ) : (
            ""
          )}
          {index != sortDpDisplayArr.length - 1 ? (
            <>
              <img
                title="Move down"
                style={{
                  cursor:
                    dpUpdate && dpData.length == dpDisplayData.length
                      ? "pointer"
                      : "default",
                }}
                src={`${MoveDown}`}
                width={36}
                height={36}
                onClick={(_) => {
                  dpUpdate && dpData.length == dpDisplayData.length
                    ? dpArrows(index, "Down")
                    : null;
                }}
              />
            </>
          ) : (
            ""
          )}
          {item.Source != "DP" && item.ID != 0 ? (
            <>
              <Icon
                iconName="Delete"
                style={{
                  paddingLeft:
                    index == 0 || index == dpDisplayData.length - 1 ? 35 : 0,
                }}
                title="Delete deliverable"
                className={dpiconStyleClass.delete}
                onClick={() => {
                  setdpButtonLoader(false),
                    setdpDeletePopup({ condition: true, targetId: item.ID });
                }}
              />
            </>
          ) : (
            ""
          )}
        </div>
      ),
    },
  ];
  const dpDrpDwnOptns = {
    source: [{ key: "All", text: "All" }],
    status: [{ key: "All", text: "All" }],
    developer: [{ key: "All", text: "All" }],
  };
  const dpFilterKeys = {
    source: "All",
    status: "All",
    developer: "All",
  };
  let dpErrorStatus = {
    Deliverable: "",
    Source: "",
  };
  const dpAddItems = {
    RefId: 0,
    ID: 0,
    AnnualPlanID: Ap_AnnualPlanId,
    Source: "CIM",
    ProductId: null,
    Title: "",
    NotApplicable: false,
    NotApplicableManager: false,
    IsCompleteStatus: false,
    StartDate: null,
    EndDate: null,
    Status: "Scheduled",
    DeveloperId: null,
    ManagerId: null,
    PlannedHours: 0,
    Week: Dp_WeekNumber,
    Year: Dp_Year,
    TLink: "",
    PBLink: "",
    DateError: false,
    BA: null,
    ActualHours: 0,
    Onchange: true,
    IsNew: true,
    IsCompleteNew: false,
    Developer: null,
  };
  const Source = [
    { key: "CIM", text: "CIM" },
    { key: "OM", text: "OM" },
  ];
  // Design
  const dpProjectInfo = mergeStyles({
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    margin: "0 15px 0 0",
  });
  const dplabelStyles = mergeStyleSets({
    titleLabel: [
      {
        color: "#676767",
        fontSize: "14px",
        marginRight: "10px",
        fontWeight: "400",
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
        marginTop: "25px",
        marginLeft: "10px",
        fontWeight: "500",
        color: "#323130",
        fontSize: "13px",
      },
    ],
  });

  const dpdropdownStyles: Partial<IDropdownStyles> = {
    root: { width: 186, marginRight: 15 },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      border: "1px solid #E8E8EA",
    },
    dropdownItemsWrapper: { backgroundColor: "#F5F5F7", fontSize: 12 },
    dropdownItemSelected: { backgroundColor: "#DCDCDC", fontSize: 12 },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const dpiconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const dpiconStyleClass = mergeStyleSets({
    link: [{ color: "blue", margin: "0 0" }, dpiconStyle],
    delete: [{ color: "red", margin: "0 7px" }, dpiconStyle],
    edit: [{ color: "blue", margin: "0 7px" }, dpiconStyle],
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 22,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginTop: 28,
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
  });
  const dpBigiconStyle = mergeStyles({
    fontSize: 25,
    height: 20,
    width: 25,
    cursor: "pointer",
    marginRight: 10,
    marginTop: 2,
  });
  const dpBigiconStyleClass = mergeStyleSets({
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
  const dpprofileName = mergeStyles({
    fontWeight: "bold",
    marginRight: "10px",
  });

  // detailslist
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
        },
      },
    },

    headerWrapper: {
      flex: "0 0 auto",
    },
    contentWrapper: {
      // flex: "1 1 auto",
      // overflow: "hidden",
      height: "calc(100vh - 345px)",
      overflowX: "hidden",
      overflowY: "auto",
    },
  };
  const gridStyles_Nodata: Partial<IDetailsListStyles> = {
    root: {
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          ".ms-DetailsRow-fields": {
            alignItems: "center",
            height: 38,
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
      // height: "calc(100vh - 345px)",
      overflowX: "hidden",
      overflowY: "auto",
    },
  };

  const dpstatusStyle = mergeStyles({
    textAlign: "center",
    paddingTop: 2,
    borderRadius: "25px",
  });
  const dpstatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      dpstatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      dpstatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      dpstatusStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      dpstatusStyle,
    ],
    Onhold: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      dpstatusStyle,
    ],
  });
  const dpbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const dpbuttonStyleClass = mergeStyleSets({
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
      dpbuttonStyle,
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
      dpbuttonStyle,
    ],
  });
  const dpTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: { width: 400, marginLeft: 25, marginTop: 15 },
    field: { backgroundColor: "whitesmoke", fontSize: 12 },
  };
  // Heading Styles
  const dpTxtHeadingBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 450,
      margin: 0,
      backgroundColor: "#2392B2",
      textAlign: "center",
      height: 40,
      padding: 10,
      fontSize: 18,
      fontWeight: 600,
      color: "White",
    },
  };

  // useState
  const [dpIsCompleted, setdpIsCompleted] = useState(false);
  const [dpNACheckbox, setdpNACheckbox] = useState(false);
  const [dpNAMCheckbox, setdpNAMCheckbox] = useState(false);
  const [dpUpdate, setdpUpdate] = useState(false);
  const [dpReRender, setdpReRender] = useState(false);
  const [apCurrentData, setapCurrentData] = useState([]);
  const [dpData, setdpData] = useState(_dpAllitems);
  const [dpMasterData, setdpMasterData] = useState(_dpAllitems);
  const [dpDisplayData, setdpDisplayData] = useState(_dpAllitems);
  const [dpDeliverable, setdpDeliverable] = useState(dpAddItems);
  const [dpDropDownOptions, setdpDropDownOptions] = useState(dpDrpDwnOptns);
  const [dpFilterOptions, setdpFilterOptions] = useState(dpFilterKeys);
  const [dpModalBoxVisibility, setdpModalBoxVisibility] = useState(false);
  const [dpShowMessage, setdpShowMessage] = useState(dpErrorStatus);
  const [dpLoader, setdpLoader] = useState(true);
  const [dpErrorDate, setdpErrorDate] = useState(false);
  const [dpButtonLoader, setdpButtonLoader] = useState(false);
  const [dpAutoSave, setdpAutoSave] = useState(false);
  const [thisweekPBData, setthisweekPBData] = useState([]);
  const [dpColumns, setdpColumns] = useState(_dpColumns);
  const [dpDeletePopup, setdpDeletePopup] = useState({
    condition: false,
    targetId: 0,
  });

  const [dpConfirmationPopup, setdpConfirmationPopup] = useState({
    condition: false,
  });

  const dateFormater = (date: Date): string => {
    return !date ? "" : moment(date).format("DD/MM/YYYY");
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

  // PlayBook
  const getDeliveryPlanTemplate = (_apData) => {
    sharepointWeb.lists
      .getByTitle("Delivery Plan Phase List")
      .items.select("*", "Phases/Title", "Phases/Id")
      .expand("Phases")
      // .filter("DeliverPlanTypeOfWork eq '" + _apData.TypeofProject + "' ")
      .top(5000)
      .get()
      .then((items) => {
        gblDeliveryPlanTemplate = items;
      })
      .catch((err) => {
        dpErrorFunction(err, "getDeliveryPlanTemplate");
      });
  };

  // Getting data from Delivery Plan List and Dropdown Options
  const getthisweekPBData = () => {
    sharepointWeb.lists
      .getByTitle("ProductionBoard")
      .items.filter(
        "Week eq '" +
          Dp_WeekNumber +
          "' and Year eq '" +
          Dp_Year +
          "' and AnnualPlanID eq '" +
          Ap_AnnualPlanId +
          "' "
      )
      .top(5000)
      .get()
      .then(async (items) => {
        setthisweekPBData([...items]);
      })
      .catch((err) => {
        dpErrorFunction(err, "getthisweekPBData");
      });
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
        "Master_x0020_Project/ProductVersion",
        "Master_x0020_Project/Title",
        "Master_x0020_Project/Id",
        "FieldValuesAsText/StartDate",
        "FieldValuesAsText/PlannedEndDate"
      )
      .expand(
        "ProjectOwner",
        "ProjectLead",
        "Master_x0020_Project",
        "FieldValuesAsText"
      )
      .filter("ID eq '" + Ap_AnnualPlanId + "' ")
      .top(5000)
      .get()
      .then(async (items) => {
        items.forEach((item) => {
          _apCurrentData.push({
            ID: item.ID,
            Title: item.Title,
            ProjectVersion: item.ProjectVersion ? item.ProjectVersion : "V1",
            TypeofProject: item.ProjectType,
            ProductId: item.Master_x0020_ProjectId,
            ProductName: item.Master_x0020_Project
              ? item.Master_x0020_Project.Title
              : "",
            ProductVersion: item.Master_x0020_Project
              ? item.Master_x0020_Project.ProductVersion
                ? item.Master_x0020_Project.ProductVersion
                : "V1"
              : "V1",
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
            StartDate: moment(
              item["FieldValuesAsText"].StartDate,
              DateListFormat
            ).format(DateListFormat),
            PlannedEndDate: moment(
              item["FieldValuesAsText"].PlannedEndDate,
              DateListFormat
            ).format(DateListFormat),
            AllocatedHours: item.AllocatedHours ? item.AllocatedHours : 0,
            BA: item.BusinessArea,
            MultipleDeveloper: item.ProjectLead ? item.ProjectLead : null,
            Status: item.Status ? item.Status : null,
          });
        });

        setapCurrentData([..._apCurrentData]);
        getDpData(_apCurrentData[0]);
        getDeliveryPlanTemplate(_apCurrentData[0]);
      })
      .catch((err) => {
        dpErrorFunction(err, "getCurrentAPData");
      });
  };
  const getDpData = (_apData) => {
    sharepointWeb.lists
      .getByTitle("Delivery Plan")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,Manager/Title,Manager/EMail,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
      )
      .expand("Developer, Manager,FieldValuesAsText")
      .filter("AnnualPlanID eq '" + Ap_AnnualPlanId + "' ")
      .orderBy("OrderId", true)
      .top(5000)
      .get()
      .then((items) => {
        items.forEach((item: any, index: number) => {
          _dpAllitems.push({
            RefId: index + 1,
            ID: item.ID,
            AnnualPlanID: Ap_AnnualPlanId,
            Source: item.Source,
            ProductId: item.ProductId,
            Title: item.Title,
            NotApplicable: item.NotApplicable,
            NotApplicableManager: item.NotApplicableManager,
            IsCompleteStatus: item.Status == "Completed" ? true : false,
            StartDate: item.StartDate
              ? moment(
                  item["FieldValuesAsText"].StartDate,
                  DateListFormat
                ).format(DateListFormat)
              : null,
            EndDate: item.EndDate
              ? moment(
                  item["FieldValuesAsText"].EndDate,
                  DateListFormat
                ).format(DateListFormat)
              : null,
            Status: item.Status,
            DeveloperId: item.DeveloperId,
            ManagerId: item.ManagerId ? item.ManagerId : null,
            PlannedHours: item.PlannedHours ? item.PlannedHours : 0,
            Week: item.Week,
            Year: item.Year,
            TLink: "",
            PBLink: props.playbookURL + Ap_AnnualPlanId,
            DateError: false,
            BA: item.BA,
            ActualHours: item.ActualHours ? item.ActualHours : 0,
            Onchange: false,
            IsNew: false,
            IsCompleteNew: false,
            Developer: item.DeveloperId
              ? allPeoples.filter((people) => {
                  return people.ID == item.DeveloperId;
                })[0].text
              : null,
          });
        });

        if (_dpAllitems.length == 0) {
          getDpTemplateData(_apData);
        } else {
          console.log(_dpAllitems);
          setdpDropDownOptions(dpDrpDwnOptns);
          setdpMasterData([..._dpAllitems]);
          setdpData([..._dpAllitems]);
          sortDpDataArr = _dpAllitems;
          setdpDisplayData([..._dpAllitems]);
          sortDpDisplayArr = _dpAllitems;
          reloadFilterOptions(_dpAllitems);
          setdpColumns(_dpColumns);
          setdpLoader(false);
        }
      })
      .catch((err) => {
        dpErrorFunction(err, "getDpData");
      });
  };
  const getDpTemplateData = (_apData) => {
    sharepointWeb.lists
      .getByTitle("Delivery Plan Phase List")
      .items.filter("DeliverPlanTypeOfWork eq '" + _apData.TypeofProject + "' ")
      .top(5000)
      .get()
      .then((items) => {
        items.forEach((item: any, index: number) => {
          _dpAllitems.push({
            RefId: index + 1,
            ID: 0,
            AnnualPlanID: Ap_AnnualPlanId,
            Source: "DP",
            ProductId: _apData.ProductId,
            Title: item.Title,
            NotApplicable: false,
            NotApplicableManager: false,
            IsCompleteStatus: false,
            StartDate: _apData.StartDate ? _apData.StartDate : null,
            EndDate: _apData.PlannedEndDate ? _apData.PlannedEndDate : null,
            Status: "Scheduled",
            DeveloperId: _apData.DeveloperId ? _apData.DeveloperId : null,
            ManagerId: _apData.ProjectOwnerId ? _apData.ProjectOwnerId : null,
            PlannedHours: item.Hours ? item.Hours : 0,
            Week: Dp_WeekNumber,
            Year: Dp_Year,
            TLink: "",
            PBLink: "",
            DateError: false,
            BA: _apData.BA,
            ActualHours: 0,
            Onchange: false,
            IsNew: false,
            IsCompleteNew: false,
            Developer: _apData.DeveloperId
              ? allPeoples.filter((people) => {
                  return people.ID == _apData.DeveloperId;
                })[0].text
              : null,
          });
        });
        if (_dpAllitems.length > 0) {
          setdpDropDownOptions(dpDrpDwnOptns);
          setdpMasterData([..._dpAllitems]);
          setdpData([..._dpAllitems]);
          sortDpDataArr = _dpAllitems;
          setdpDisplayData([..._dpAllitems]);
          sortDpDisplayArr = _dpAllitems;
          reloadFilterOptions(_dpAllitems);
          setdpColumns(_dpColumns);
          setdpLoader(false);
        } else {
          setdpLoader(false);
        }
      })
      .catch((err) => {
        dpErrorFunction(err, "getDpTemplateData");
      });
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;

    const sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    tempArrReload.forEach((dp) => {
      if (
        dpDrpDwnOptns.source.findIndex((source) => {
          return source.key == dp.Source;
        }) == -1
      ) {
        dpDrpDwnOptns.source.push({
          key: dp.Source,
          text: dp.Source,
        });
      }
      if (
        dpDrpDwnOptns.status.findIndex((status) => {
          return status.key == dp.Status;
        }) == -1
      ) {
        dpDrpDwnOptns.status.push({
          key: dp.Status,
          text: dp.Status,
        });
      }
      if (
        dpDrpDwnOptns.developer.findIndex((developer) => {
          return developer.key == dp.DeveloperId;
        }) == -1 &&
        dp.DeveloperId
      ) {
        dpDrpDwnOptns.developer.push({
          key: dp.DeveloperId,
          text: dp.Developer,
        });
      }
    });

    if (
      dpDrpDwnOptns.developer.some((managerOptn) => {
        return (
          managerOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      dpDrpDwnOptns.developer.shift();
      let loginUserIndex = dpDrpDwnOptns.developer.findIndex((user) => {
        return (
          user.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      });
      let loginUserData = dpDrpDwnOptns.developer.splice(loginUserIndex, 1);

      dpDrpDwnOptns.developer.sort(sortFilterKeys);
      dpDrpDwnOptns.developer.unshift(loginUserData[0]);
      dpDrpDwnOptns.developer = usersOrderFunction(dpDrpDwnOptns.developer);
      dpDrpDwnOptns.developer.unshift({ key: "All", text: "All" });
    } else {
      dpDrpDwnOptns.developer.shift();
      dpDrpDwnOptns.developer.sort(sortFilterKeys);
      dpDrpDwnOptns.developer = usersOrderFunction(dpDrpDwnOptns.developer);
      dpDrpDwnOptns.developer.unshift({ key: "All", text: "All" });
    }

    setdpLoader(false);
    setdpDropDownOptions(dpDrpDwnOptns);
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
  // Button Click Function
  const cancelDPData = () => {
    setdpFilterOptions({ ...dpFilterKeys });
    reloadFilterOptions(dpMasterData);
    setdpDisplayData([...dpMasterData]);
    sortDpDisplayArr = dpMasterData;
    setdpData([...dpMasterData]);
    sortDpDataArr = dpMasterData;
    setdpUpdate(false);
    sortDpUpdate = false;
    setdpNACheckbox(false);
    setdpNAMCheckbox(false);
  };
  const saveDPData = () => {
    setdpLoader(true);
    let Hours = sumOfHours();
    let successCount = 0;

    let apCompletedValue = dpData.filter((arr) => {
      return arr.IsCompleteStatus == true;
    });

    let apStatusValue =
      dpData.length == apCompletedValue.length
        ? "Completed"
        : apCurrentData[0].Status;
    apCurrentData[0].Status == apStatusValue;
    setapCurrentData([...apCurrentData]);

    dpData.forEach((dp, Index: number) => {
      let strDSNA: string = `${dp.DeveloperId}-${
        dp.Status == "Completed" ? 1 : 0
      }-${
        dp.NotApplicable != true && dp.NotApplicableManager != true
          ? false
          : true
      }`;
      let statusValue = dp.IsCompleteStatus
        ? "Completed"
        : dp.Status
        ? dp.Status
        : null;

      let requestdata = {
        AnnualPlanIDId: dp.AnnualPlanID ? dp.AnnualPlanID : Ap_AnnualPlanId,
        Source: dp.Source ? dp.Source : null,
        Title: dp.Title ? dp.Title : null,
        StartDate: dp.StartDate
          ? moment(dp.StartDate, DateListFormat).format("YYYY-MM-DD")
          : null,
        EndDate: dp.EndDate
          ? moment(dp.EndDate, DateListFormat).format("YYYY-MM-DD")
          : null,
        DeveloperId: dp.DeveloperId ? dp.DeveloperId : null,
        ManagerId: dp.ManagerId ? dp.ManagerId : null,
        NotApplicable: dp.NotApplicable ? dp.NotApplicable : null,
        NotApplicableManager: dp.NotApplicableManager
          ? dp.NotApplicableManager
          : null,
        Status: statusValue,
        PlannedHours: dp.PlannedHours ? dp.PlannedHours : 0,
        ProductId: apCurrentData.length > 0 ? apCurrentData[0].ProductId : null,
        Week: Dp_WeekNumber,
        Year: Dp_Year,
        OrderId: Index,
        BA: dp.BA ? dp.BA : null,
        AnnualPlanIDNumber: dp.AnnualPlanID ? dp.AnnualPlanID : Ap_AnnualPlanId,
        Project: apCurrentData.length > 0 ? apCurrentData[0].Title : "",
        ProjectVersion:
          apCurrentData.length > 0 ? apCurrentData[0].ProjectVersion : "V1",
        ActualHours: dp.ActualHours ? dp.ActualHours : 0,
        SPFxFilter: strDSNA,
      };
      if (dp.ID != 0) {
        sharepointWeb.lists
          .getByTitle("Delivery Plan")
          .items.getById(dp.ID)
          .update(requestdata)
          .then((e) => {
            successCount++;
            dpData[Index].Status = statusValue;
            dpData[Index].IsCompleteNew = false;
            let disIndex = dpDisplayData.findIndex(
              (obj) => obj.RefId == dp.RefId
            );
            dpDisplayData[disIndex].Status = statusValue;

            let updatePB = thisweekPBData.filter((pb) => {
              return pb.DeliveryPlanID == dp.ID;
            });

            if (updatePB.length > 0 && dp.Onchange == true) {
              let strDWYNA: string = `${
                dp.DeveloperId
              }-${Dp_WeekNumber}-${Dp_Year}-${
                dp.NotApplicable != true && dp.NotApplicableManager != true
                  ? false
                  : true
              }`;
              sharepointWeb.lists
                .getByTitle("ProductionBoard")
                .items.getById(updatePB[0].ID)
                .update({
                  ProductId:
                    apCurrentData.length > 0
                      ? apCurrentData[0].ProductId
                      : null,
                  StartDate: dp.StartDate
                    ? moment(dp.StartDate, DateListFormat).format("YYYY-MM-DD")
                    : null,
                  EndDate: dp.EndDate
                    ? moment(dp.EndDate, DateListFormat).format("YYYY-MM-DD")
                    : null,
                  DeveloperId: dp.DeveloperId ? dp.DeveloperId : null,
                  NotApplicable: dp.NotApplicable,
                  NotApplicableManager: dp.NotApplicableManager,
                  PlannedHours: dp.PlannedHours ? dp.PlannedHours : 0,
                  SPFxFilter: strDWYNA,
                })
                .then((e) => {})
                .catch((err) => {
                  dpErrorFunction(err, "saveDPData-getPBItem");
                });
            }

            if (dpData.length == successCount) {
              sharepointWeb.lists
                .getByTitle(ListNameURL)
                .items.getById(Ap_AnnualPlanId)
                .update({ Status: apStatusValue, AllocatedHours: Hours })
                .then((e) => {
                  apCurrentData[0].AllocatedHours = Hours;
                })
                .catch((err) => {
                  dpErrorFunction(err, "saveDPData-getAPItem");
                });

              setdpDisplayData([...dpDisplayData]);
              sortDpDisplayArr = dpDisplayData;

              setdpData([...dpData]);
              sortDpDataArr = dpData;
              setdpMasterData([...dpData]);
              setdpUpdate(false);
              sortDpUpdate = false;

              setdpNACheckbox(false);
              setdpNAMCheckbox(false);

              setdpColumns(_dpColumns);
              AddSuccessPopup();
              setdpLoader(false);
            }
          })
          .catch((err) => {
            dpErrorFunction(err, "saveDPData-updateDPItem");
          });
      } else if (dp.ID == 0) {
        sharepointWeb.lists
          .getByTitle("Delivery Plan")
          .items.add(requestdata)
          .then((e) => {
            successCount++;
            dpData[Index].ID = e.data.ID;
            dpData[Index].Status = statusValue;
            dpData[Index].IsCompleteNew = false;
            dpData[Index].PBLink = props.playbookURL + Ap_AnnualPlanId;

            let disIndex = dpDisplayData.findIndex(
              (obj) => obj.RefId == dp.RefId
            );
            dpDisplayData[disIndex].Status = statusValue;

            if (thisweekPBData.length > 0 && dp.Source != "DP") {
              let strDWYNA: string = `${
                dp.DeveloperId
              }-${Dp_WeekNumber}-${Dp_Year}-${
                dp.NotApplicable != true && dp.NotApplicableManager != true
                  ? false
                  : true
              }`;
              sharepointWeb.lists
                .getByTitle("ProductionBoard")
                .items.add({
                  BA: dp.BA ? dp.BA : null,
                  StartDate: dp.StartDate
                    ? moment(dp.StartDate, DateListFormat).format("YYYY-MM-DD")
                    : null,
                  EndDate: dp.EndDate
                    ? moment(dp.EndDate, DateListFormat).format("YYYY-MM-DD")
                    : null,
                  Source: dp.Source ? dp.Source : null,
                  AnnualPlanIDId: dp.AnnualPlanID ? dp.AnnualPlanID : null,
                  ProductId:
                    apCurrentData.length > 0
                      ? apCurrentData[0].ProductId
                      : null,
                  Title: dp.Title ? dp.Title : null,
                  PlannedHours: dp.PlannedHours ? dp.PlannedHours : null,
                  Monday: "0",
                  Tuesday: "0",
                  Wednesday: "0",
                  Thursday: "0",
                  Friday: "0",
                  ActualHours: 0,
                  DeveloperId: dp.DeveloperId ? dp.DeveloperId : null,
                  Week: Dp_WeekNumber,
                  Year: Dp_Year,
                  NotApplicable: dp.NotApplicable,
                  NotApplicableManager: dp.NotApplicableManager,
                  DeliveryPlanID: dp.ID,
                  DPActualHours: 0,
                  // Status: "Pending",
                  AnnualPlanIDNumber: dp.AnnualPlanID,
                  Project:
                    apCurrentData.length > 0 ? apCurrentData[0].Title : "",
                  ProjectVersion:
                    apCurrentData.length > 0
                      ? apCurrentData[0].ProjectVersion
                      : "V1",
                  SPFxFilter: strDWYNA,
                })
                .then((e) => {})
                .catch((err) => {
                  dpErrorFunction(err, "saveDPData-addPBItem");
                });
            }

            if (dpData.length == successCount) {
              sharepointWeb.lists
                .getByTitle(ListNameURL)
                .items.getById(Ap_AnnualPlanId)
                .update({ Status: apStatusValue, AllocatedHours: Hours })
                .then((e) => {
                  apCurrentData[0].AllocatedHours = Hours;
                })
                .catch((err) => {
                  dpErrorFunction(err, "saveDPData-addAPItem");
                });

              setdpDisplayData([...dpDisplayData]);
              sortDpDisplayArr = dpDisplayData;

              setdpData([...dpData]);
              sortDpDataArr = dpData;
              setdpMasterData([...dpData]);

              setdpUpdate(false);
              sortDpUpdate = false;

              setdpNACheckbox(false);
              setdpNAMCheckbox(false);

              setdpColumns(_dpColumns);
              AddSuccessPopup();
              setdpLoader(false);
            }
          })
          .catch((err) => {
            dpErrorFunction(err, "saveDPData-addDPItem");
          });
      }
    });
  };
  const dpDeleteItem = (id: number) => {
    sharepointWeb.lists
      .getByTitle("Delivery Plan")
      .items.getById(id)
      .delete()
      .then(() => {
        let tempPBArr = thisweekPBData.filter((arr) => {
          return arr.DeliveryPlanID == id;
        });
        if (tempPBArr.length > 0) {
          sharepointWeb.lists
            .getByTitle("ProductionBoard")
            .items.getById(tempPBArr[0].ID)
            .delete()
            .then(() => {
              let tempMasterArr = [...dpMasterData];
              let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
              tempMasterArr.splice(targetIndex, 1);

              let temp_ap_arr = [...dpData];
              let targetIndexapdata = temp_ap_arr.findIndex(
                (arr) => arr.ID == id
              );
              temp_ap_arr.splice(targetIndexapdata, 1);

              setdpMasterData([...tempMasterArr]);
              setdpData([...temp_ap_arr]);
              sortDpDataArr = temp_ap_arr;
              setdpDisplayData([...temp_ap_arr]);
              sortDpDisplayArr = temp_ap_arr;
              setdpLoader(true);
              reloadFilterOptions(tempMasterArr);

              setdpUpdate(false);
              sortDpUpdate = false;
              setdpColumns([..._dpColumns]);
              setdpDeletePopup({ condition: false, targetId: 0 });
              DeleteSuccessPopup();
            })
            .catch((err) => {
              dpErrorFunction(err, "dpDeleteItem-deletePBItem");
            });
        } else {
          let tempMasterArr = [...dpMasterData];
          let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
          tempMasterArr.splice(targetIndex, 1);

          let temp_ap_arr = [...dpData];
          let targetIndexapdata = temp_ap_arr.findIndex((arr) => arr.ID == id);
          temp_ap_arr.splice(targetIndexapdata, 1);

          setdpMasterData([...tempMasterArr]);
          setdpData([...temp_ap_arr]);
          sortDpDataArr = temp_ap_arr;
          setdpDisplayData([...temp_ap_arr]);
          sortDpDisplayArr = temp_ap_arr;
          setdpLoader(true);
          reloadFilterOptions(tempMasterArr);

          setdpUpdate(false);
          sortDpUpdate = false;
          setdpColumns([..._dpColumns]);
          setdpDeletePopup({ condition: false, targetId: 0 });
          DeleteSuccessPopup();
        }
      })
      .catch((err) => {
        dpErrorFunction(err, "dpDeleteItem-deleteDPItem");
      });
  };

  // Onchange and Filters
  const dpListFilter = (key, option) => {
    let tempArr = [...dpData];
    let tempDpFilterKeys = { ...dpFilterOptions };
    tempDpFilterKeys[`${key}`] = option;

    if (tempDpFilterKeys.source != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Source == tempDpFilterKeys.source;
      });
    }
    if (tempDpFilterKeys.status != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Status == tempDpFilterKeys.status;
      });
    }
    if (tempDpFilterKeys.developer != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.DeveloperId == tempDpFilterKeys.developer;
      });
    }

    // setdpData([...tempArr]);
    setdpDisplayData([...tempArr]);
    sortDpDisplayArr = tempArr;
    setdpFilterOptions({ ...tempDpFilterKeys });
  };
  const dpAddOnchange = (key, value) => {
    let tempArronchange = dpDeliverable;
    if (key == "deliverable") dpDeliverable.Title = value;
    else if (key == "source") dpDeliverable.Source = value;

    dpDeliverable.ManagerId = apCurrentData[0].ProjectOwnerId;
    dpDeliverable.DeveloperId = apCurrentData[0].DeveloperId;
    dpDeliverable.ProductId = apCurrentData[0].ProductId;
    dpDeliverable.BA = apCurrentData[0].BA;
    dpDeliverable.EndDate = apCurrentData[0].PlannedEndDate
      ? apCurrentData[0].PlannedEndDate
      : new Date();
    dpDeliverable.StartDate = apCurrentData[0].StartDate
      ? apCurrentData[0].StartDate
      : new Date();
    dpDeliverable.Developer = apCurrentData[0].DeveloperId
      ? allPeoples.filter((people) => {
          return people.ID == apCurrentData[0].DeveloperId;
        })[0].text
      : null;
    dpDeliverable.RefId = dpData.length + 1;
    setdpDeliverable(tempArronchange);
  };
  const dpArrows = (Index, Type) => {
    if (Type == "Up") {
      var srcUp = dpData[Index];
      var desUp = dpData[Index - 1];

      dpData[Index] = desUp;
      dpData[Index - 1] = srcUp;

      setdpDisplayData([...dpData]);
      sortDpDisplayArr = dpData;
      setdpData([...dpData]);
      sortDpDataArr = dpData;
    }
    if (Type == "Down") {
      var srcDown = dpData[Index];
      var desDown = dpData[Index + 1];

      dpData[Index] = desDown;
      dpData[Index + 1] = srcDown;

      setdpDisplayData([...dpData]);
      sortDpDisplayArr = dpData;
      setdpData([...dpData]);
      sortDpDataArr = dpData;
    }
  };
  const dpOnchangeItems = (RefId, key, value) => {
    let Index = dpData.findIndex((obj) => obj.RefId == RefId);
    let disIndex = dpDisplayData.findIndex((obj) => obj.RefId == RefId);
    let dpBeforeData = dpData[Index];
    let dpOnchangeData = [
      {
        RefId: dpBeforeData.RefId,
        ID: dpBeforeData.ID,
        AnnualPlanID: dpBeforeData.AnnualPlanID,
        Source: dpBeforeData.Source,
        ProductId: dpBeforeData.ProductId,
        Title: dpBeforeData.Title,
        NotApplicable:
          key == "NotApplicable" ? value : dpBeforeData.NotApplicable,
        NotApplicableManager:
          key == "NotApplicableManager"
            ? value
            : dpBeforeData.NotApplicableManager,
        IsCompleteStatus:
          key == "IsCompleteStatus" ? value : dpBeforeData.IsCompleteStatus,
        StartDate: key == "StartDate" ? value : dpBeforeData.StartDate,
        EndDate: key == "EndDate" ? value : dpBeforeData.EndDate,
        Status: dpBeforeData.Status,
        DeveloperId: key == "DeveloperId" ? value : dpBeforeData.DeveloperId,
        ManagerId: dpBeforeData.ManagerId,
        PlannedHours: key == "PlannedHours" ? value : dpBeforeData.PlannedHours,
        Week: dpBeforeData.Week,
        Year: dpBeforeData.Year,
        TLink: dpBeforeData.TLink,
        PBLink: dpBeforeData.PBLink,
        DateError: dpBeforeData.DateError,
        BA: dpBeforeData.BA,
        ActualHours: dpBeforeData.ActualHours,
        Onchange: true,
        IsNew: dpBeforeData.IsNew,
        IsCompleteNew:
          key == "IsCompleteStatus" ? value : dpBeforeData.IsCompleteNew,
        Developer:
          key == "DeveloperId" && value
            ? allPeoples.filter((people) => {
                return people.ID == value;
              })[0].text
            : dpBeforeData.DeveloperId
            ? allPeoples.filter((people) => {
                return people.ID == dpBeforeData.DeveloperId;
              })[0].text
            : null,
      },
    ];

    dpData[Index] = dpOnchangeData[0];
    dpDisplayData[disIndex] = dpOnchangeData[0];

    // Common checkbox values

    let isNADetails = dpData.filter((arr) => {
      return arr.NotApplicable == true;
    });
    let isNAMDetails = dpData.filter((arr) => {
      return arr.NotApplicableManager == true;
    });
    let isCompleteDetails = dpData.filter((arr) => {
      return arr.IsCompleteStatus == true;
    });

    dpData.length == isNADetails.length
      ? setdpNACheckbox(true)
      : setdpNACheckbox(false);
    dpData.length == isNAMDetails.length
      ? setdpNAMCheckbox(true)
      : setdpNAMCheckbox(false);
    dpData.length == isCompleteDetails.length
      ? setdpIsCompleted(true)
      : setdpIsCompleted(false);

    reloadFilterOptions(dpData);
    setdpData([...dpData]);
    sortDpDataArr = dpData;
  };

  // Header Content
  const sumOfHours = () => {
    let WNAdpData =
      dpData.length > 0
        ? dpData.filter((arr) => {
            return (
              arr.NotApplicable != true && arr.NotApplicableManager != true
            );
          })
        : [];
    var sum: number = 0;
    if (WNAdpData.length > 0) {
      WNAdpData.forEach((x) => {
        sum += parseFloat(x.PlannedHours ? x.PlannedHours : 0);
      });
      return sum % 1 == 0 ? sum : sum.toFixed(2);
    } else {
      return 0;
    }
  };
  const sumOfActualHours = () => {
    var sum: number = 0;
    if (dpData.length > 0) {
      dpData.forEach((x) => {
        sum += parseFloat(x.ActualHours ? x.ActualHours : 0);
      });
      return sum % 1 == 0 ? sum : sum.toFixed(2);
    } else {
      return 0;
    }
  };
  const overallStatus = () => {
    let curData = dpData.filter(
      (dp) => dp.NotApplicable != true && dp.NotApplicableManager != true
    );
    if (curData.every((data) => data.Status == "Completed")) {
      return (
        <div
          style={{ width: "125px" }}
          className={dpstatusStyleClass.completed}
        >
          Completed
        </div>
      );
    } else if (curData.every((data) => data.Status == "Scheduled")) {
      return (
        <div
          style={{ width: "125px" }}
          className={dpstatusStyleClass.scheduled}
        >
          Scheduled
        </div>
      );
    } else if (curData.every((data) => data.Status == "On schedule")) {
      return (
        <div
          style={{ width: "125px" }}
          className={dpstatusStyleClass.onSchedule}
        >
          On schedule
        </div>
      );
    } else if (curData.every((data) => data.Status == "Behind schedule")) {
      return (
        <div
          style={{ width: "125px" }}
          className={dpstatusStyleClass.behindScheduled}
        >
          Behind schedule
        </div>
      );
    } else if (curData.every((data) => data.Status == "On hold")) {
      return (
        <div style={{ width: "125px" }} className={dpstatusStyleClass.Onhold}>
          On hold
        </div>
      );
    } else {
      return (
        <div
          style={{ width: "125px" }}
          className={dpstatusStyleClass.scheduled}
        >
          Scheduled
        </div>
      );
    }
  };

  // Validation and Success
  const dpValidationFunction = () => {
    let isError = false;

    let errorStatus = {
      Deliverable: "",
      Source: "",
    };

    if (!dpDeliverable.Title) {
      isError = true;
      errorStatus.Deliverable = "Please add deliverable";
    }
    if (!dpDeliverable.Source) {
      isError = true;
      errorStatus.Source = "Please Select a value for Source";
    }

    if (!isError) {
      setdpButtonLoader(true);
      setdpDisplayData(dpDisplayData.concat(dpDeliverable));
      setdpData(dpData.concat(dpDeliverable));
      setdpModalBoxVisibility(false);
      reloadFilterOptions(dpDisplayData.concat(dpDeliverable));
      setdpUpdate(true);
      console.log(dpData.concat(dpDeliverable));

      //Sorting
      sortDpUpdate = true;
      sortDpDisplayArr = dpDisplayData.concat(dpDeliverable);
      sortDpDataArr = dpData.concat(dpDeliverable);
      setdpColumns(_dpColumns);
    } else {
      setdpShowMessage(errorStatus);
    }
  };
  const dpErrorFunction = (error: any, functionName: string) => {
    console.log(error, functionName);
    let response = {
      ComponentName: "Delivery plan",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setdpLoader(false);
        setdpButtonLoader(false);
        if (dpData.length > 0) {
          ErrorPopup();
        }
      }
    );
  };

  const AddSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Delivery plan is successfully submitted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const DeleteSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Delivery plan is successfully deleted !!!")
  );
  const dpDateErrorFunction = () => {
    let resultDateError = dpData.filter((dp) => dp.DateError == true);

    if (resultDateError.length > 0) {
      setdpErrorDate(true);
    } else {
      setdpErrorDate(false);
    }
  };

  window.onbeforeunload = function (e) {
    debugger;
    if (dpAutoSave) {
      let dialogText =
        "You have unsaved changes, are you sure you want to leave?";
      e.returnValue = dialogText;
      return dialogText;
    }
  };

  const alertDialog = () => {
    if (confirm("You have unsaved changes, are you sure you want to leave?")) {
      props.handleclick("AnnualPlan", null, "DP");
    }
  };

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _dpColumns;
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

    const newDpDataArr = _copyAndSort(
      sortDpDataArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newDPDisplayArr = _copyAndSort(
      sortDpDisplayArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setdpData([...newDpDataArr]);
    setdpDisplayData([...newDPDisplayArr]);
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

  useEffect(() => {
    if (dpAutoSave && !dpErrorDate && dpUpdate) {
      setTimeout(() => {
        document.getElementById("dpBtnSave").click();
      }, 300000);
    }
  }, [dpAutoSave]);

  useEffect(() => {
    getCurrentAPData();
    getthisweekPBData();
  }, [dpReRender]);

  // return function are dispaly in page
  return (
    <>
      {dpLoader ? (
        <CustomLoader />
      ) : (
        <div style={{ padding: "5px 15px" }}>
          {/* {dpLoader ? <CustomLoader /> : null} */}
          <div
            className={styles.dpHeaderSection}
            style={{ paddingBottom: "0 " }}
          >
            <div
              style={{
                // position: "sticky",
                // top: 0,
                // backgroundColor: "#fff",
                // zIndex: 1,
                marginBottom: 41,
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "space-between",
                  marginBottom: 20,
                  color: "#2392b2",
                }}
              >
                {/* Header Start */}
                <div className={styles.dpTitle}>
                  <Icon
                    aria-label="ChevronLeftMed"
                    iconName="NavigateBack"
                    className={dpBigiconStyleClass.ChevronLeftMed}
                    onClick={() => {
                      dpAutoSave
                        ? confirm(
                            "You have unsaved changes, are you sure you want to leave?"
                          )
                          ? props.handleclick("AnnualPlan", null, "DP")
                          : null
                        : props.handleclick("AnnualPlan", null, "DP");
                    }}
                  />
                  <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
                    Delivery plan
                  </Label>
                </div>
                <div style={{ display: "flex" }}>
                  {
                    <div className={styles.userDetails}>
                      <Label
                        className={dpprofileName}
                        style={{ display: "flex" }}
                      >
                        <div style={{ color: "#E2A00F", width: 50 }}>
                          Client:{" "}
                        </div>
                        {apCurrentData.length > 0
                          ? apCurrentData[0].ProjectOwnerName
                          : ""}
                      </Label>
                      <Persona
                        size={PersonaSize.size24}
                        presence={PersonaPresence.none}
                        imageUrl={
                          "/_layouts/15/userphoto.aspx?size=S&username=" +
                          `${
                            apCurrentData.length > 0
                              ? apCurrentData[0].ProjectOwnerEmail
                              : ""
                          }`
                        }
                      />
                    </div>
                  }
                  {
                    <div className={styles.userDetails}>
                      <Label
                        className={dpprofileName}
                        style={{ display: "flex" }}
                      >
                        <div style={{ color: "#E2A00F", width: 80 }}>
                          Developer:{" "}
                        </div>
                        {apCurrentData.length > 0
                          ? apCurrentData[0].DeveloperName
                          : ""}
                      </Label>
                      <Persona
                        size={PersonaSize.size24}
                        presence={PersonaPresence.none}
                        imageUrl={
                          "/_layouts/15/userphoto.aspx?size=S&username=" +
                          `${
                            apCurrentData.length > 0
                              ? apCurrentData[0].DeveloperEmail
                              : ""
                          }`
                        }
                      />
                      {apCurrentData.length > 1 ? (
                        <TooltipHost
                          content={
                            <ul style={{ margin: 10, padding: 0 }}>
                              {apCurrentData[0].MultipleDeveloper.map(
                                (DName) => {
                                  return (
                                    <li>
                                      <div style={{ display: "flex" }}>
                                        <Persona
                                          showOverflowTooltip
                                          size={PersonaSize.size24}
                                          presence={PersonaPresence.none}
                                          showInitialsUntilImageLoads={true}
                                          imageUrl={
                                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                                            `${DName.EMail}`
                                          }
                                        />
                                        <Label style={{ marginLeft: 10 }}>
                                          {DName.Title}
                                        </Label>
                                      </div>
                                    </li>
                                  );
                                }
                              )}
                            </ul>
                          }
                          styles={{ root: { display: "inline-block" } }}
                        >
                          <div className={styles.extraPeople}>
                            {apCurrentData[0].MultipleDeveloper.length}
                          </div>
                        </TooltipHost>
                      ) : null}
                    </div>
                  }
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                }}
              >
                <div className={styles.Section1}>
                  {apCurrentData.length > 0 &&
                  apCurrentData[0].TypeofProject != null ? (
                    <PrimaryButton
                      text="Add deliverable"
                      className={dpbuttonStyleClass.buttonPrimary}
                      onClick={(_) => {
                        setdpDeliverable(dpAddItems);
                        setdpShowMessage(dpErrorStatus);
                        setdpModalBoxVisibility(true);
                        setdpButtonLoader(false);
                      }}
                    />
                  ) : (
                    <PrimaryButton
                      text="Add deliverable"
                      disabled={true}
                      onClick={(_) => {
                        // setdpDeliverable(dpAddItems);
                        // setdpShowMessage(dpErrorStatus);
                        // setdpModalBoxVisibility(true);
                        // setdpButtonLoader(false);
                      }}
                    />
                  )}
                </div>
                {dpData.length > 0 ? (
                  <div>
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                      }}
                    >
                      {dpUpdate ? (
                        <PrimaryButton
                          iconProps={cancelIcon}
                          text="Cancel"
                          className={dpbuttonStyleClass.buttonPrimary}
                          onClick={(_) => {
                            setdpAutoSave(false);
                            setdpErrorDate(false);
                            cancelDPData();
                          }}
                        />
                      ) : (
                        <PrimaryButton
                          iconProps={editIcon}
                          text="Edit"
                          className={dpbuttonStyleClass.buttonPrimary}
                          onClick={(_) => {
                            setdpAutoSave(true);
                            setdpUpdate(true);
                            setdpNACheckbox(false);
                            setdpNAMCheckbox(false);
                            //Sorting
                            sortDpUpdate = true;
                            setdpColumns(_dpColumns);
                            setdpData([...dpMasterData]);

                            let tempArr = [...dpMasterData];
                            let tempDpFilterKeys = { ...dpFilterOptions };
                            if (tempDpFilterKeys.source != "All") {
                              tempArr = tempArr.filter((arr) => {
                                return arr.Source == tempDpFilterKeys.source;
                              });
                            }
                            if (tempDpFilterKeys.status != "All") {
                              tempArr = tempArr.filter((arr) => {
                                return arr.Status == tempDpFilterKeys.status;
                              });
                            }
                            if (tempDpFilterKeys.developer != "All") {
                              tempArr = tempArr.filter((arr) => {
                                return (
                                  arr.DeveloperId == tempDpFilterKeys.developer
                                );
                              });
                            }

                            setdpDisplayData([...tempArr]);
                          }}
                        />
                      )}

                      {dpErrorDate || !dpUpdate ? (
                        <PrimaryButton
                          iconProps={saveIcon}
                          text="Save"
                          disabled={true}
                          onClick={(_) => {
                            // setdpAutoSave(false);
                            // saveDPData();
                          }}
                        />
                      ) : (
                        <PrimaryButton
                          iconProps={saveIcon}
                          id="dpBtnSave"
                          text="Save"
                          className={dpbuttonStyleClass.buttonSecondary}
                          onClick={(_) => {
                            setdpAutoSave(false);

                            let isCompletedData = dpData.filter((arr) => {
                              return arr.IsCompleteNew == true;
                            });
                            if (isCompletedData.length > 0) {
                              setdpConfirmationPopup({
                                condition: true,
                              });
                            } else {
                              saveDPData();
                            }
                          }}
                        />
                      )}
                      <Icon
                        title="Production board"
                        iconName="Link12"
                        className={dpiconStyleClass.pblink}
                        onClick={() => {
                          dpAutoSave
                            ? confirm(
                                "You have unsaved changes, are you sure you want to leave?"
                              )
                              ? props.handleclick(
                                  "ProductionBoard",
                                  Ap_AnnualPlanId,
                                  "DP"
                                )
                              : null
                            : props.handleclick(
                                "ProductionBoard",
                                Ap_AnnualPlanId,
                                "DP"
                              );
                        }}
                      />
                    </div>
                  </div>
                ) : null}
              </div>
              <div
                style={{
                  display: "flex",
                  marginTop: 15,
                  justifyContent: "space-between",
                }}
              >
                <div className={styles.Section1}>
                  <div className={dpProjectInfo}>
                    <Label className={dplabelStyles.titleLabel}>
                      Name of the deliverable :
                    </Label>
                    <Label
                      className={dplabelStyles.labelValue}
                      style={{ maxWidth: 250 }}
                    >
                      {apCurrentData.length > 0
                        ? apCurrentData[0].Title +
                          " " +
                          apCurrentData[0].ProjectVersion
                        : ""}
                    </Label>
                  </div>
                  <div className={dpProjectInfo}>
                    <Label className={dplabelStyles.titleLabel}>
                      Product :
                    </Label>
                    <Label
                      className={dplabelStyles.labelValue}
                      style={{ maxWidth: 250 }}
                    >
                      {apCurrentData.length > 0
                        ? apCurrentData[0].ProductName +
                          " " +
                          apCurrentData[0].ProductVersion
                        : ""}
                    </Label>
                  </div>
                  <div className={dpProjectInfo}>
                    <Label className={dplabelStyles.titleLabel}>
                      Actual hrs/ Planned hrs :
                    </Label>
                    <Label className={dplabelStyles.labelValue}>
                      {sumOfActualHours()} / {sumOfHours()}
                    </Label>
                  </div>
                  <div className={dpProjectInfo}>
                    <Label className={dplabelStyles.titleLabel}>Status :</Label>
                    {dpData.length > 0 ? overallStatus() : ""}
                  </div>
                  <div className={dpProjectInfo}>
                    <Label className={dplabelStyles.titleLabel}>TOD :</Label>
                    <Label className={dplabelStyles.labelValue}>
                      {apCurrentData.length > 0
                        ? apCurrentData[0].TypeofProject
                        : null}
                    </Label>
                  </div>
                </div>
              </div>

              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  marginBottom: "15px",
                  paddingBottom: "10px",
                  flexWrap: "wrap",
                }}
              >
                <div className={styles.ddSection}>
                  {/* Section Start */}
                  <div>
                    <Label className={dplabelStyles.inputLabels}>Source</Label>
                    <Dropdown
                      selectedKey={dpFilterOptions.source}
                      placeholder="Select an option"
                      options={dpDropDownOptions.source}
                      styles={dpdropdownStyles}
                      onChange={(e, option: any) => {
                        dpListFilter("source", option["key"]);
                      }}
                    />
                  </div>

                  <div>
                    <Label className={dplabelStyles.inputLabels}>Status</Label>
                    <Dropdown
                      selectedKey={dpFilterOptions.status}
                      placeholder="Select an option"
                      options={dpDropDownOptions.status}
                      styles={dpdropdownStyles}
                      onChange={(e, option: any) => {
                        dpListFilter("status", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label className={dplabelStyles.inputLabels}>
                      Developer
                    </Label>
                    <Dropdown
                      selectedKey={dpFilterOptions.developer}
                      placeholder="Select an option"
                      options={dpDropDownOptions.developer}
                      styles={dpdropdownStyles}
                      onChange={(e, option: any) => {
                        dpListFilter("developer", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label className={dplabelStyles.inputLabels}>N/A</Label>
                    <Checkbox
                      styles={{
                        root: { marginTop: 3, width: 50 },
                      }}
                      disabled={!dpUpdate ? true : false}
                      checked={dpNACheckbox}
                      onChange={(ev) => {
                        setdpNACheckbox(!dpNACheckbox);
                        dpData.forEach((item, Index) => {
                          let dpBeforeData = dpData[Index];
                          let dpOnchangeData = [
                            {
                              RefId: dpBeforeData.RefId,
                              ID: dpBeforeData.ID,
                              AnnualPlanID: dpBeforeData.AnnualPlanID,
                              Source: dpBeforeData.Source,
                              ProductId: dpBeforeData.ProductId,
                              Title: dpBeforeData.Title,
                              NotApplicable: ev.target["checked"],
                              NotApplicableManager:
                                dpBeforeData.NotApplicableManager,
                              IsCompleteStatus: dpBeforeData.IsCompleteStatus,
                              StartDate: dpBeforeData.StartDate,
                              EndDate: dpBeforeData.EndDate,
                              Status: dpBeforeData.Status,
                              DeveloperId: dpBeforeData.DeveloperId,
                              ManagerId: dpBeforeData.ManagerId,
                              PlannedHours: dpBeforeData.PlannedHours,
                              Week: dpBeforeData.Week,
                              Year: dpBeforeData.Year,
                              TLink: dpBeforeData.TLink,
                              PBLink: dpBeforeData.PBLink,
                              DateError: dpBeforeData.DateError,
                              BA: dpBeforeData.BA,
                              ActualHours: dpBeforeData.ActualHours,
                              Onchange: true,
                              IsNew: dpBeforeData.IsNew,
                              IsCompleteNew: dpBeforeData.IsCompleteNew,
                              Developer: dpBeforeData.Developer,
                            },
                          ];
                          dpData[Index] = dpOnchangeData[0];
                        });

                        dpDisplayData.forEach((item, Index) => {
                          let dpBeforeData = dpDisplayData[Index];
                          let dpOnchangeData = [
                            {
                              RefId: dpBeforeData.RefId,
                              ID: dpBeforeData.ID,
                              AnnualPlanID: dpBeforeData.AnnualPlanID,
                              Source: dpBeforeData.Source,
                              ProductId: dpBeforeData.ProductId,
                              Title: dpBeforeData.Title,
                              NotApplicable: ev.target["checked"],
                              NotApplicableManager:
                                dpBeforeData.NotApplicableManager,
                              IsCompleteStatus: dpBeforeData.IsCompleteStatus,
                              StartDate: dpBeforeData.StartDate,
                              EndDate: dpBeforeData.EndDate,
                              Status: dpBeforeData.Status,
                              DeveloperId: dpBeforeData.DeveloperId,
                              ManagerId: dpBeforeData.ManagerId,
                              PlannedHours: dpBeforeData.PlannedHours,
                              Week: dpBeforeData.Week,
                              Year: dpBeforeData.Year,
                              TLink: dpBeforeData.TLink,
                              PBLink: dpBeforeData.PBLink,
                              DateError: dpBeforeData.DateError,
                              BA: dpBeforeData.BA,
                              ActualHours: dpBeforeData.ActualHours,
                              Onchange: true,
                              IsNew: dpBeforeData.IsNew,
                              IsCompleteNew: dpBeforeData.IsCompleteNew,
                              Developer: dpBeforeData.Developer,
                            },
                          ];
                          dpDisplayData[Index] = dpOnchangeData[0];
                        });

                        setdpData([...dpData]);
                        sortDpDataArr = dpData;
                      }}
                    />
                  </div>
                  <div>
                    <Label className={dplabelStyles.inputLabels}>N/A(C)</Label>
                    <Checkbox
                      styles={{
                        root: { marginTop: 3, width: 50 },
                      }}
                      disabled={
                        dpUpdate &&
                        (apCurrentData.length > 0 && loggeduseremail != ""
                          ? apCurrentData[0].ProjectOwnerEmail ==
                            loggeduseremail
                          : null)
                          ? false
                          : true
                      }
                      checked={dpNAMCheckbox}
                      onChange={(ev) => {
                        setdpNAMCheckbox(!dpNAMCheckbox);
                        dpData.forEach((item, Index) => {
                          let dpBeforeData = dpData[Index];
                          let dpOnchangeData = [
                            {
                              RefId: dpBeforeData.RefId,
                              ID: dpBeforeData.ID,
                              AnnualPlanID: dpBeforeData.AnnualPlanID,
                              Source: dpBeforeData.Source,
                              ProductId: dpBeforeData.ProductId,
                              Title: dpBeforeData.Title,
                              NotApplicable: dpBeforeData.NotApplicable,
                              NotApplicableManager: ev.target["checked"],
                              IsCompleteStatus: dpBeforeData.IsCompleteStatus,
                              StartDate: dpBeforeData.StartDate,
                              EndDate: dpBeforeData.EndDate,
                              Status: dpBeforeData.Status,
                              DeveloperId: dpBeforeData.DeveloperId,
                              ManagerId: dpBeforeData.ManagerId,
                              PlannedHours: dpBeforeData.PlannedHours,
                              Week: dpBeforeData.Week,
                              Year: dpBeforeData.Year,
                              TLink: dpBeforeData.TLink,
                              PBLink: dpBeforeData.PBLink,
                              DateError: dpBeforeData.DateError,
                              BA: dpBeforeData.BA,
                              ActualHours: dpBeforeData.ActualHours,
                              Onchange: true,
                              IsNew: dpBeforeData.IsNew,
                              IsCompleteNew: dpBeforeData.IsCompleteNew,
                              Developer: dpBeforeData.Developer,
                            },
                          ];
                          dpData[Index] = dpOnchangeData[0];
                        });

                        dpDisplayData.forEach((item, Index) => {
                          let dpBeforeData = dpDisplayData[Index];
                          let dpOnchangeData = [
                            {
                              RefId: dpBeforeData.RefId,
                              ID: dpBeforeData.ID,
                              AnnualPlanID: dpBeforeData.AnnualPlanID,
                              Source: dpBeforeData.Source,
                              ProductId: dpBeforeData.ProductId,
                              Title: dpBeforeData.Title,
                              NotApplicable: dpBeforeData.NotApplicable,
                              NotApplicableManager: ev.target["checked"],
                              IsCompleteStatus: dpBeforeData.IsCompleteStatus,
                              StartDate: dpBeforeData.StartDate,
                              EndDate: dpBeforeData.EndDate,
                              Status: dpBeforeData.Status,
                              DeveloperId: dpBeforeData.DeveloperId,
                              ManagerId: dpBeforeData.ManagerId,
                              PlannedHours: dpBeforeData.PlannedHours,
                              Week: dpBeforeData.Week,
                              Year: dpBeforeData.Year,
                              TLink: dpBeforeData.TLink,
                              PBLink: dpBeforeData.PBLink,
                              DateError: dpBeforeData.DateError,
                              BA: dpBeforeData.BA,
                              ActualHours: dpBeforeData.ActualHours,
                              Onchange: true,
                              IsNew: dpBeforeData.IsNew,
                              IsCompleteNew: dpBeforeData.IsCompleteNew,
                              Developer: dpBeforeData.Developer,
                            },
                          ];
                          dpDisplayData[Index] = dpOnchangeData[0];
                        });

                        setdpData([...dpData]);
                        sortDpDataArr = dpData;
                      }}
                    />
                  </div>

                  <div>
                    <Label className={dplabelStyles.inputLabels}>
                      Complete
                    </Label>
                    <Checkbox
                      styles={{
                        root: { marginTop: 3, width: 50 },
                      }}
                      disabled={!dpUpdate ? true : false}
                      checked={dpIsCompleted}
                      onChange={(ev) => {
                        setdpIsCompleted(!dpIsCompleted);
                        dpData.forEach((item, Index) => {
                          let dpBeforeData = dpData[Index];
                          let dpOnchangeData = [
                            {
                              RefId: dpBeforeData.RefId,
                              ID: dpBeforeData.ID,
                              AnnualPlanID: dpBeforeData.AnnualPlanID,
                              Source: dpBeforeData.Source,
                              ProductId: dpBeforeData.ProductId,
                              Title: dpBeforeData.Title,
                              NotApplicable: dpBeforeData.NotApplicable,
                              NotApplicableManager:
                                dpBeforeData.NotApplicableManager,
                              IsCompleteStatus:
                                item.Status != "Completed"
                                  ? ev.target["checked"]
                                  : dpBeforeData.IsCompleteStatus,
                              StartDate: dpBeforeData.StartDate,
                              EndDate: dpBeforeData.EndDate,
                              Status: dpBeforeData.Status,
                              DeveloperId: dpBeforeData.DeveloperId,
                              ManagerId: dpBeforeData.ManagerId,
                              PlannedHours: dpBeforeData.PlannedHours,
                              Week: dpBeforeData.Week,
                              Year: dpBeforeData.Year,
                              TLink: dpBeforeData.TLink,
                              PBLink: dpBeforeData.PBLink,
                              DateError: dpBeforeData.DateError,
                              BA: dpBeforeData.BA,
                              ActualHours: dpBeforeData.ActualHours,
                              Onchange: true,
                              IsNew: dpBeforeData.IsNew,
                              IsCompleteNew:
                                item.Status != "Completed"
                                  ? ev.target["checked"]
                                  : dpBeforeData.IsCompleteNew,
                              Developer: dpBeforeData.Developer,
                            },
                          ];
                          dpData[Index] = dpOnchangeData[0];
                        });

                        dpDisplayData.forEach((item, Index) => {
                          let dpBeforeData = dpDisplayData[Index];
                          let dpOnchangeData = [
                            {
                              RefId: dpBeforeData.RefId,
                              ID: dpBeforeData.ID,
                              AnnualPlanID: dpBeforeData.AnnualPlanID,
                              Source: dpBeforeData.Source,
                              ProductId: dpBeforeData.ProductId,
                              Title: dpBeforeData.Title,
                              NotApplicable: dpBeforeData.NotApplicable,
                              NotApplicableManager:
                                dpBeforeData.NotApplicableManager,
                              IsCompleteStatus:
                                item.Status != "Completed"
                                  ? ev.target["checked"]
                                  : dpBeforeData.IsCompleteStatus,
                              StartDate: dpBeforeData.StartDate,
                              EndDate: dpBeforeData.EndDate,
                              Status: dpBeforeData.Status,
                              DeveloperId: dpBeforeData.DeveloperId,
                              ManagerId: dpBeforeData.ManagerId,
                              PlannedHours: dpBeforeData.PlannedHours,
                              Week: dpBeforeData.Week,
                              Year: dpBeforeData.Year,
                              TLink: dpBeforeData.TLink,
                              PBLink: dpBeforeData.PBLink,
                              DateError: dpBeforeData.DateError,
                              BA: dpBeforeData.BA,
                              ActualHours: dpBeforeData.ActualHours,
                              Onchange: true,
                              IsNew: dpBeforeData.IsNew,
                              IsCompleteNew:
                                item.Status != "Completed"
                                  ? ev.target["checked"]
                                  : dpBeforeData.IsCompleteNew,
                              Developer: dpBeforeData.Developer,
                            },
                          ];
                          dpDisplayData[Index] = dpOnchangeData[0];
                        });

                        setdpData([...dpData]);
                        sortDpDataArr = dpData;
                      }}
                    />
                  </div>
                  <div>
                    <Icon
                      title="Click to reset"
                      iconName="Refresh"
                      className={dpiconStyleClass.refresh}
                      onClick={() => {
                        if (dpAutoSave) {
                          if (
                            confirm(
                              "You have unsaved changes, are you sure you want to leave?"
                            )
                          ) {
                            setdpData([...dpMasterData]);
                            sortDpDataArr = dpMasterData;
                            setdpDisplayData([...dpMasterData]);
                            sortDpDisplayArr = dpMasterData;
                            setdpFilterOptions({ ...dpFilterKeys });
                            setdpUpdate(false);
                            sortDpUpdate = false;
                            setdpColumns(_dpColumns);
                          }
                        } else {
                          setdpData([...dpMasterData]);
                          sortDpDataArr = dpMasterData;
                          setdpDisplayData([...dpMasterData]);
                          sortDpDisplayArr = dpMasterData;
                          setdpFilterOptions({ ...dpFilterKeys });
                          setdpUpdate(false);
                          sortDpUpdate = false;
                          setdpColumns(_dpColumns);
                        }
                      }}
                    />
                  </div>

                  {dpErrorDate ? (
                    <div>
                      <Label className={dplabelStyles.ErrorLabel}>
                        Please choose valid dates
                      </Label>
                    </div>
                  ) : null}
                  {/* Section Start */}
                </div>

                <div>
                  <Label className={dplabelStyles.NORLabel}>
                    Number of records:{" "}
                    <b style={{ color: "#038387" }}>{dpDisplayData.length}</b>
                  </Label>
                </div>
              </div>
            </div>
            {/* Header- End */}
          </div>
          <div style={{ marginTop: -40 }}>
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
            title="Go to top"
            className={styles.scrollTop}
            onClick={() => {
              document.querySelector("#forFocus")["focus"]();
            }}
          >
            <Icon
              iconName="Up"
              className={dpiconStyleClass.link}
              style={{ color: "#fff" }}
            />
          </div>
          <div>
            {
              <DetailsList
                items={dpDisplayData}
                columns={sortDpUpdate ? _dpColumns : dpColumns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                // data-is-scrollable={true}
                onShouldVirtualize={() => {
                  return false;
                }}
                styles={dpData.length == 0 ? gridStyles_Nodata : gridStyles}
                // styles={{ root: { width: "100%" } }}
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
          {dpData.length == 0 ? (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                marginTop: "15px",
              }}
            >
              <Label style={{ color: "#2392B2" }}>No data Found !!!</Label>
            </div>
          ) : null}

          <div>
            <Modal isOpen={dpModalBoxVisibility} isBlocking={false}>
              <div>
                {" "}
                <Label styles={dpTxtHeadingBoxStyles}>Add deliverable</Label>
              </div>
              <TextField
                styles={dpTxtBoxStyles}
                required={true}
                errorMessage={dpShowMessage.Deliverable}
                label="Deliverable"
                onChange={(e, value: string) => {
                  dpAddOnchange("deliverable", value);
                }}
              />
              <div>
                <ChoiceGroup
                  styles={dpTxtBoxStyles}
                  defaultSelectedKey="CIM"
                  options={Source}
                  label="Source"
                  onChange={(e, option: any) => {
                    dpAddOnchange("source", option["key"]);
                  }}
                />
              </div>
              <div></div>
              <div className={styles.apModalBoxButtonSection}>
                <button
                  className={styles.apModalBoxSubmitBtn}
                  onClick={(_) => {
                    setdpAutoSave(true);
                    dpValidationFunction();
                    document.querySelector("#forFocusBottom")["focus"]();
                  }}
                >
                  {dpButtonLoader ? (
                    <Spinner />
                  ) : (
                    <span>
                      <Icon
                        iconName="Save"
                        style={{ marginTop: 4, marginRight: 12 }}
                      />
                      {"Add"}
                    </span>
                  )}
                </button>
                <button
                  className={styles.apModalBoxBackBtn}
                  onClick={(_) => {
                    setdpShowMessage(dpErrorStatus);
                    setdpDeliverable(dpAddItems);
                    setdpModalBoxVisibility(false);
                  }}
                >
                  <span
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                    }}
                  >
                    {" "}
                    <Icon
                      iconName="ChromeBack"
                      style={{ marginTop: 4, marginRight: 12 }}
                    />
                    Close
                  </span>
                </button>
              </div>
            </Modal>
          </div>
          <div>
            <Modal isOpen={dpDeletePopup.condition} isBlocking={true}>
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
                    The data up to production board will be deleted. Are you
                    sure want to delete?
                  </Label>
                </div>
              </div>
              <div className={styles.apDeletePopupBtnSection}>
                <button
                  onClick={(_) => {
                    setdpButtonLoader(true);
                    dpDeleteItem(dpDeletePopup.targetId);
                  }}
                  className={styles.apDeletePopupYesBtn}
                >
                  {dpButtonLoader ? <Spinner /> : "Yes"}
                </button>
                <button
                  onClick={(_) => {
                    setdpDeletePopup({ condition: false, targetId: 0 });
                  }}
                  className={styles.apDeletePopupNoBtn}
                >
                  No
                </button>
              </div>
            </Modal>
          </div>

          <div>
            <Modal isOpen={dpConfirmationPopup.condition} isBlocking={true}>
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
                    Confirmation
                  </Label>
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
                    setdpConfirmationPopup({ condition: false });
                    saveDPData();
                  }}
                  className={styles.apDeletePopupYesBtn}
                >
                  Yes
                </button>
                <button
                  onClick={(_) => {
                    setdpConfirmationPopup({ condition: false });
                  }}
                  className={styles.apDeletePopupNoBtn}
                >
                  No
                </button>
              </div>
            </Modal>
          </div>

          <div>
            {/* dont remove */}
            <input
              id="forFocusBottom"
              type="text"
              style={{
                width: 0,
                height: 0,
                border: "none",
                padding: 20,
              }}
            />
          </div>
        </div>
      )}
    </>
  );
};

export default DeliveryPlan;
