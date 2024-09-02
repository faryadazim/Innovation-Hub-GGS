import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
import * as moment from "moment";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  SearchBox,
  ISearchBoxStyles,
  TooltipHost,
  TooltipOverflowMode,
  IColumn,
  ILabelStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  Modal,
  TextField,
  NormalPeoplePicker,
  ITextFieldStyles,
  Spinner,
  PrimaryButton,
} from "@fluentui/react";
import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";
import {
  arraysEqual,
  IDetailsListStyles,
  Toggle,
} from "office-ui-fabric-react";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";

interface IProps {
  context: WebPartContext;
  spcontext: any;
  graphContent: any;
  URL: string;
  handleclick: any;
  pageType: string;
  peopleList: any;
  isAdmin: boolean;
}

interface IData {
  ID: number;
  BA: string;
  Title: string;
  Frequency: string;
  Responsible: string;
  Provider: string;
  Audience: string;
  Status: string;
  ResponsibleId: number;
  ProviderId: number;
  AudienceId: number;
  ConfigID: number;
  TimePeriod: string;
  Year: number;
  TLink: string;
  DueDate: string;
}
interface IFilter {
  BA: string;
  Title: string;
  Frequency: string;
  Audience: string;
  Status: string;
}

interface IDropdowns {
  BA: [{ key: string; text: string }];
  Title: [{ key: string; text: string }];
  Frequency: [{ key: string; text: string }];
  Audience: [{ key: string; text: string }];
  Status: [{ key: string; text: string }];
}

let sortORData: IData[] = [];
let sortORFilterData: IData[] = [];
let PendingData: IData[] = [];
let _pendingColumns = [];

const OrgReporting = (props: IProps): JSX.Element => {
  const sharepointWeb: any = Web(props.URL);
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  const allPeoples: any[] = props.peopleList;
  let CurrentPage: number = 1;
  let totalPageItems: number = 10;
  let OR_Year: number = moment().year();
  let OR_WeekNumber: number = moment().isoWeek();
  let OR_Month: string = moment().format("MMMM");
  let OR_Term: string =
    moment().month() >= 10
      ? "Term 4"
      : moment().month() >= 7
      ? "Term 3"
      : moment().month() >= 4
      ? "Term 2"
      : "Term 1";

  const ORFilterKeys: IFilter = {
    BA: "All",
    Title: "All",
    Frequency: "All",
    Audience: "All",
    Status: "All",
  };

  const ORFilterOptns: IDropdowns = {
    BA: [{ key: "All", text: "All" }],
    Title: [{ key: "All", text: "All" }],
    Frequency: [{ key: "All", text: "All" }],
    Audience: [{ key: "All", text: "All" }],
    Status: [{ key: "All", text: "All" }],
  };

  let ORErrorStatus = {
    Request: "",
    Requestto: "",
    Documenttype: "",
    Link: "",
  };

  const ORModalBoxDrpDwnOptns = {
    Request: [],
    Documenttype: [],
  };

  const drAllitems = {
    Request: null,
    Requestto: null,
    Emailcc: null,
    Project: null,
    Documenttype: null,
    Link: null,
    Comments: null,
    Confidential: false,
    Product: null,
    AnnualPlanID: null,
    DeliveryPlanID: null,
    ProductionBoardID: null,
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

  const _ORColumns: IColumn[] = [
    {
      key: "Column1",
      name: "Business area",
      fieldName: "BA",
      minWidth: 120,
      maxWidth: 120,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) =>
        BAacronymsCollection.filter((ba) => {
          return ba.Name == item.BA;
        })[0].ShortName,
    },
    {
      key: "Column2",
      name: "Title",
      fieldName: "Title",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
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
      key: "Column3",
      name: "Frequency",
      fieldName: "Frequency",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column4",
      name: "Due date",
      fieldName: "DueDate",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        let frequencyType =
          item.Frequency == "Weekly"
            ? "LDW"
            : item.Frequency == "Monthly"
            ? "LDM"
            : item.Frequency == "Term"
            ? "LDT"
            : "";
        let typeAbbreviations =
          item.Frequency == "Weekly"
            ? "Last day of week"
            : item.Frequency == "Monthly"
            ? "Last day of month"
            : item.Frequency == "Term"
            ? "Last day of term"
            : "";
        return (
          <>
            <div
              style={{
                marginTop: "-6px",
              }}
              title={
                moment(item.DueDate).format("DD/MM/YYYY") +
                ` ( ${typeAbbreviations} )`
              }
            >
              {moment(item.DueDate).format("DD/MM/YYYY") +
                ` ( ${frequencyType} )`}
            </div>
          </>
        );
      },
    },
    // {
    //   key: "Column4",
    //   name: "Responsible",
    //   fieldName: "Responsible",
    //   minWidth: 100,
    //   maxWidth: 200,
    //   onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
    //     _onColumnClick(ev, column);
    //   },
    //   onRender: (item) => (
    //     <div style={{ display: "flex" }}>
    //       <div
    //         style={{
    //           marginTop: "-6px",
    //         }}
    //         title={item.Responsible}
    //       >
    //         <Persona
    //           size={PersonaSize.size32}
    //           presence={PersonaPresence.none}
    //           imageUrl={
    //             "/_layouts/15/userphoto.aspx?size=S&username=" +
    //             `${
    //                allPeoples.filter((ap) => {
    //   return ap.ID == item.ResponsibleId;
    // }).length > 0
    //   ? allPeoples.filter((ap) => {
    //       return ap.ID == item.ResponsibleId;
    //     })[0].secondaryText
    //   : null
    //             }`
    //           }
    //         />
    //       </div>
    //       <div>
    //         <span title={item.Responsible} style={{ fontSize: "13px" }}>
    //           {item.Responsible}
    //         </span>
    //       </div>
    //     </div>
    //   ),
    // },
    {
      key: "Column5",
      name: "Provider",
      fieldName: "Provider",
      minWidth: 100,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <div style={{ display: "flex" }}>
          <div
            style={{
              marginTop: "-6px",
            }}
            title={item.Provider}
          >
            <Persona
              size={PersonaSize.size32}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${
                  allPeoples.filter((ap) => {
                    return ap.ID == item.ProviderId;
                  }).length > 0
                    ? allPeoples.filter((ap) => {
                        return ap.ID == item.ProviderId;
                      })[0].secondaryText
                    : null
                }`
              }
            />
          </div>
          <div>
            <span title={item.Provider} style={{ fontSize: "13px" }}>
              {item.Provider}
            </span>
          </div>
        </div>
      ),
    },
    {
      key: "Column6",
      name: "Audience",
      fieldName: "Audience",
      minWidth: 100,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <div style={{ display: "flex" }}>
          <div
            style={{
              marginTop: "-6px",
            }}
            title={item.Audience}
          >
            <Persona
              size={PersonaSize.size32}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${
                  allPeoples.filter((ap) => {
                    return ap.ID == item.AudienceId;
                  }).length > 0
                    ? allPeoples.filter((ap) => {
                        return ap.ID == item.AudienceId;
                      })[0].secondaryText
                    : null
                }`
              }
            />
          </div>
          <div>
            <span title={item.Audience} style={{ fontSize: "13px" }}>
              {item.Audience}
            </span>
          </div>
        </div>
      ),
    },
    {
      key: "Column7",
      name: "Status",
      fieldName: "Status",
      minWidth: 100,
      maxWidth: 250,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          {item.Status == "Read" ? (
            <div className={ORStatusStyleClass.completed}>{item.Status}</div>
          ) : item.Status == "Scheduled" ? (
            <div className={ORStatusStyleClass.scheduled}>{item.Status}</div>
          ) : item.Status == "Submitted" ? (
            <div className={ORStatusStyleClass.submitted}>{item.Status}</div>
          ) : (
            item.Status
          )}
        </>
      ),
    },
    {
      key: "Column11",
      name: "TL",
      fieldName: "TLink",
      minWidth: 30,
      maxWidth: 30,
      onRender: (item) => (
        <>
          {item.TLink ? (
            <a target="_blank" href={item.TLink}>
              <Icon
                iconName="NavigateExternalInline"
                className={ORiconStyleClass.link}
                style={{ color: "#038387" }}
              />
            </a>
          ) : (
            <Icon
              iconName="NavigateExternalInline"
              className={ORiconStyleClass.link}
              style={{ color: "#038387" }}
            />
          )}
        </>
      ),
    },
    {
      key: "Column8",
      name: "Action",
      fieldName: "Action",
      minWidth: 75,
      maxWidth: 150,
      onRender: (item, Index) => (
        <>
          {item.ID ? (
            <Icon
              iconName="OpenEnrollment"
              style={{
                color:
                  item.Status == "Scheduled"
                    ? "#0882A5"
                    : item.Status == "Read"
                    ? "#40b200"
                    : item.Status == "Submitted"
                    ? "#B3B300"
                    : "#000000",
                marginLeft: 9,
                fontSize: 17,
                height: 14,
                width: 17,
                cursor: "pointer",
              }}
              onClick={(_) => {
                drAllitems.Project = item.Title;
                drAllitems.ProductionBoardID = item.ID;
                setORButtonLoader(false);
                setORShowMessage(ORErrorStatus);
                setORDocumentReview(drAllitems);
                setORModalBoxVisibility(true);
              }}
            />
          ) : (
            <Icon
              iconName="OpenEnrollment"
              style={{
                color: "#ababab",
                marginLeft: 9,
                fontSize: 17,
                height: 14,
                width: 17,
                cursor: "default",
              }}
              onClick={(_) => {}}
            />
          )}
          {PendingData.length > 0 &&
          PendingData.filter((Dt) => {
            return Dt.ConfigID == item.ConfigID;
          }).length > 0 ? (
            <div
              style={{
                // position: "absolute",
                // z-index: 2,
                background: "#038387",
                color: "#fff",
                width: 20,
                height: 20,
                display: "inline-flex",
                alignItems: "center",
                justifyContent: "center",
                borderRadius: "50%",
                marginLeft: 5,
                cursor: "pointer",
                marginTop: 4,
              }}
              aria-describedby={item.ID}
              onClick={(_) => {
                _pendingColumns = [
                  {
                    key: "column1",
                    //name: "Time Period",
                    name:
                      item.Frequency == "Weekly"
                        ? "Week"
                        : item.Frequency == "Monthly"
                        ? "Month"
                        : item.Frequency == "Term"
                        ? "Term"
                        : "",
                    fieldName: "TimePeriod",
                    minWidth: 150,
                    maxWidth: 200,
                  },
                  {
                    key: "column2",
                    name: "Year",
                    fieldName: "Year",
                    minWidth: 150,
                    maxWidth: 200,
                  },
                  {
                    key: "Status",
                    name: "Status",
                    fieldName: "Status",
                    minWidth: 250,
                    maxWidth: 300,
                    onRender: (item) => (
                      <>
                        {item.Status == "Over due" ? (
                          <div
                            style={{
                              width: 200,
                            }}
                            className={ORStatusStyleClass.overdue}
                          >
                            {item.Status}
                          </div>
                        ) : (
                          item.Status
                        )}
                      </>
                    ),
                  },
                  {
                    key: "column3",
                    name: "Action",
                    fieldName: "Action",
                    minWidth: 70,
                    maxWidth: 150,
                    onRender: (item) => (
                      <Icon
                        iconName="OpenEnrollment"
                        style={{
                          color:
                            item.Status == "Scheduled"
                              ? "#0882A5"
                              : item.Status == "Read"
                              ? "#40b200"
                              : item.Status == "Submitted"
                              ? "#B3B300"
                              : "#000000",
                          marginTop: 6,
                          marginLeft: 9,
                          fontSize: 17,
                          height: 14,
                          width: 17,
                          cursor: "pointer",
                        }}
                        onClick={(_) => {
                          drAllitems.Project = item.Title;
                          drAllitems.ProductionBoardID = item.ID;
                          setORButtonLoader(false);
                          setORShowMessage(ORErrorStatus);
                          setORDocumentReview(drAllitems);
                          setORModalBoxVisibility(true);
                        }}
                      />
                    ),
                  },
                ];
                setORPendingPopup({
                  condition: true,
                  selectedItem: sortORData.filter((Dt) => {
                    return Dt.ID == item.ID;
                  }),
                  pendingItem: PendingData.filter((Dt) => {
                    return Dt.ConfigID == item.ConfigID;
                  }),
                });
              }}
            >
              {
                PendingData.filter((Dt) => {
                  return Dt.ConfigID == item.ConfigID;
                }).length
              }
            </div>
          ) : null}
        </>
      ),
    },
  ];
  const ORbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const ORbuttonStyleClass = mergeStyleSets({
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
      ORbuttonStyle,
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
      ORbuttonStyle,
    ],
  });
  const ORDropdownStyles: Partial<IDropdownStyles> = {
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
  const ORActiveDropdownStyles: Partial<IDropdownStyles> = {
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
      fontSize: 10,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 10,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const ORiconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const ORiconStyleClass = mergeStyleSets({
    link: [{ color: "blue", margin: "0 0" }, ORiconStyle],
    delete: [{ color: "red", margin: "0 7px" }, ORiconStyle],
    edit: [{ color: "blue", margin: "0 7px" }, ORiconStyle],
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
  const ORlabelStyles = mergeStyleSets({
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
  const ORLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const ORModalBoxDrpDwnCalloutStyles: Partial<IDropdownStyles> = {
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
  const ORMultiTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "640px",
      margin: "10px 20px",
      borderRadius: "4px",
    },
    field: { fontSize: 12, color: "#000" },
  };
  const ORModalBoxPP = mergeStyles({
    width: "300px",
    margin: "10px 20px",
  });
  const ORTxtBoxStyles: Partial<ITextFieldStyles> = {
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
  const ORModalBoxDetailsListStyles: Partial<IDetailsListStyles> = {
    root: {
      width: 960,
      overflowX: "none",
      selectors: {
        ".ms-DetailsRow-cell": {
          height: 45,
          display: "flex",
          alignItems: "center",
          // justifyContent: "center",
        },
        ".ms-DetailsHeader-cellTitle": {
          // justifyContent: "center !important",
        },
      },
    },
    headerWrapper: {},
    contentWrapper: { height: 140, overflowX: "hidden", overflowY: "auto" },
  };
  const ORStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    width: "160px",
  });
  const ORStatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      ORStatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      ORStatusStyle,
    ],
    submitted: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      ORStatusStyle,
    ],
    overdue: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      ORStatusStyle,
    ],
    pending: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      ORStatusStyle,
    ],
  });

  // UseState
  const [ORReRender, setORReRender] = useState<boolean>(false);
  const [ORMasterData, setORMasterData] = useState<IData[]>([]);
  const [ORData, setORData] = useState<IData[]>([]);
  const [ORFilterData, setORFilterData] = useState<IData[]>([]);
  const [ORDisplayData, setORDisplayData] = useState<IData[]>([]);
  const [ORLoader, setORLoader] = useState<boolean>(false);
  const [ORFilter, setORFilter] = useState<IFilter>(ORFilterKeys);
  const [ORFilterDrpDown, setORFilterDrpDown] =
    useState<IDropdowns>(ORFilterOptns);
  const [ORModalBoxVisibility, setORModalBoxVisibility] = useState(false);
  const [ORShowMessage, setORShowMessage] = useState(ORErrorStatus);
  const [ORModalBoxDropDownOptions, setORModalBoxDropDownOptions] = useState(
    ORModalBoxDrpDwnOptns
  );
  const [ORDocumentReview, setORDocumentReview] = useState(drAllitems);
  const [ORButtonLoader, setORButtonLoader] = useState<Boolean>(false);
  const [ORColumns, setORColumns] = useState(_ORColumns);
  const [ORcurrentPage, setORCurrentPage] = useState<number>(CurrentPage);
  const [ORPendingData, setORPendingData] = useState<IData[]>([]);
  const [ORPendingPopup, setORPendingPopup] = useState({
    condition: false,
    selectedItem: [],
    pendingItem: [],
  });

  //Function
  const getOrgReporting = (): void => {
    let _ORdata: IData[] = [];
    sharepointWeb.lists
      .getByTitle("OrgReporting")
      .items.select(
        "*",
        "Responsible/Title",
        "Responsible/Id",
        "Responsible/EMail",
        "Provider/Title",
        "Provider/Id",
        "Provider/EMail",
        "Audience/Title",
        "Audience/Id",
        "Audience/EMail"
      )
      .expand("Responsible", "Provider", "Audience")
      .filter(
        "TimePeriod eq '" +
          OR_Term +
          "' or TimePeriod eq '" +
          OR_Month +
          "' or TimePeriod eq '" +
          OR_WeekNumber +
          "' and Year eq '" +
          OR_Year +
          "'"
      )
      .top(5000)
      .orderBy("Modified", false)
      .get()
      .then(async (items) => {
        let curUserId = props.peopleList.filter((arr) => {
          return arr.secondaryText == loggeduseremail;
        })[0].ID;
        items.forEach((item) => {
          if (
            props.isAdmin ||
            (!props.isAdmin &&
              (item.ResponsibleId == curUserId || item.ProviderId == curUserId))
          ) {
            _ORdata.push({
              ID: item.ID,
              BA: item.BA,
              Title: item.Title,
              Frequency: item.Frequency,
              Responsible: item.Responsible ? item.Responsible.Title : "",
              Provider: item.Provider ? item.Provider.Title : "",
              Audience: item.Audience ? item.Audience.Title : "",
              Status: item.Status,
              ResponsibleId: item.ResponsibleId,
              ProviderId: item.ProviderId,
              AudienceId: item.AudienceId,
              ConfigID: item.ConfigID,
              TimePeriod: item.TimePeriod,
              Year: item.Year,
              TLink: item.TLink,
              DueDate: item.DueDate,
            });
          }
        });
        setORFilterData([..._ORdata]);
        sortORFilterData = _ORdata;
        setORData([..._ORdata]);
        sortORData = _ORdata;
        setORMasterData([..._ORdata]);
        reloadFilterDropdowns([..._ORdata]);
        paginateFunction(1, _ORdata);
        setORLoader(false);
      })
      .catch((error) => {
        ORErrorFunction(error, "getOrgReporting");
      });
  };
  const getPendingReport = (): void => {
    let _ORpendingdata: IData[] = [];
    sharepointWeb.lists
      .getByTitle("OrgReporting")
      .items.select(
        "*",
        "Responsible/Title",
        "Responsible/Id",
        "Responsible/EMail",
        "Provider/Title",
        "Provider/Id",
        "Provider/EMail",
        "Audience/Title",
        "Audience/Id",
        "Audience/EMail"
      )
      .expand("Responsible", "Provider", "Audience")
      .filter(
        " (TimePeriod ne '" +
          OR_Term +
          "' and Frequency eq 'Term') or (TimePeriod ne '" +
          OR_Month +
          "'and Frequency eq 'Monthly') or (TimePeriod ne '" +
          OR_WeekNumber +
          "' and Frequency eq 'Weekly') and Status ne 'Read' "
      )
      .top(5000)
      .get()
      .then(async (items) => {
        console.log(items);
        items.forEach((item) => {
          _ORpendingdata.push({
            ID: item.ID,
            BA: item.BA,
            Title: item.Title,
            Frequency: item.Frequency,
            Responsible: item.Responsible ? item.Responsible.Title : "",
            Provider: item.Provider ? item.Provider.Title : "",
            Audience: item.Audience ? item.Audience.Title : "",
            Status: item.Status,
            ResponsibleId: item.ResponsibleId,
            ProviderId: item.ProviderId,
            AudienceId: item.AudienceId,
            ConfigID: item.ConfigID,
            TimePeriod: item.TimePeriod,
            Year: item.Year,
            TLink: item.TLink,
            DueDate: item.DueDate,
          });
        });
        setORPendingData([..._ORpendingdata]);
        PendingData = _ORpendingdata;
        getOrgReporting();
      })
      .catch((error) => {
        ORErrorFunction(error, "getPendingReport");
      });
  };
  const reloadFilterDropdowns = (data: IData[]): void => {
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

    tempArrReload.forEach((item) => {
      if (
        ORFilterOptns.BA.findIndex((BA) => {
          return BA.key == item.BA;
        }) == -1 &&
        item.BA
      ) {
        ORFilterDrpDown.BA.push({
          key: item.BA,
          text: item.BA,
        });
      }
      if (
        ORFilterOptns.Title.findIndex((Title) => {
          return Title.key == item.Title;
        }) == -1 &&
        item.Title
      ) {
        ORFilterDrpDown.Title.push({
          key: item.Title,
          text: item.Title,
        });
      }
      if (
        ORFilterOptns.Frequency.findIndex((Frequency) => {
          return Frequency.key == item.Frequency;
        }) == -1 &&
        item.Frequency
      ) {
        ORFilterDrpDown.Frequency.push({
          key: item.Frequency,
          text: item.Frequency,
        });
      }
      if (
        ORFilterOptns.Audience.findIndex((Audience) => {
          return Audience.key == item.Audience;
        }) == -1 &&
        item.Audience
      ) {
        ORFilterDrpDown.Audience.push({
          key: item.Audience,
          text: item.Audience,
        });
      }
      if (
        ORFilterOptns.Status.findIndex((Status) => {
          return Status.key == item.Status;
        }) == -1 &&
        item.Status
      ) {
        ORFilterDrpDown.Status.push({
          key: item.Status,
          text: item.Status,
        });
      }
    });

    if (
      ORFilterOptns.Audience.some((managerOptn) => {
        return (
          managerOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      ORFilterOptns.Audience.shift();
      let loginUserIndex = ORFilterOptns.Audience.findIndex((user) => {
        return (
          user.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      });
      let loginUserData = ORFilterOptns.Audience.splice(loginUserIndex, 1);

      ORFilterOptns.Audience.sort(sortFilterKeys);
      ORFilterOptns.Audience.unshift(loginUserData[0]);
      ORFilterOptns.Audience.unshift({ key: "All", text: "All" });
    } else {
      ORFilterOptns.Audience.shift();
      ORFilterOptns.Audience.sort(sortFilterKeys);
      ORFilterOptns.Audience.unshift({ key: "All", text: "All" });
    }

    setORFilterDrpDown(ORFilterOptns);
  };
  const getModalBoxOptions = () => {
    //Request Choices
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .fields.getByInternalNameOrTitle("Request")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              ORModalBoxDrpDwnOptns.Request.findIndex((rpb) => {
                return rpb.key == choice;
              }) == -1
            ) {
              ORModalBoxDrpDwnOptns.Request.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch(ORErrorFunction);

    //Documenttype Choices
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .fields.getByInternalNameOrTitle("Documenttype")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              ORModalBoxDrpDwnOptns.Documenttype.findIndex((rdt) => {
                return rdt.key == choice;
              }) == -1
            ) {
              ORModalBoxDrpDwnOptns.Documenttype.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch(ORErrorFunction);

    setORModalBoxDropDownOptions(ORModalBoxDrpDwnOptns);
  };
  const drValidationFunction = () => {
    let isError = false;

    let errorStatus = {
      Request: "",
      Requestto: "",
      Documenttype: "",
      Link: "",
    };

    if (!ORDocumentReview.Request) {
      isError = true;
      errorStatus.Request = "Please select a value for request";
    }
    if (!ORDocumentReview.Requestto) {
      isError = true;
      errorStatus.Requestto = "Please select a value for request to";
    }
    if (!ORDocumentReview.Documenttype) {
      isError = true;
      errorStatus.Documenttype = "Please select a value for document type";
    }
    if (!ORDocumentReview.Link) {
      isError = true;
      errorStatus.Link = "Please enter a value for link";
    }

    if (!isError) {
      setORButtonLoader(true);
      savePBDRData();
    } else {
      setORShowMessage(errorStatus);
    }
  };
  const savePBDRData = () => {
    let requestdata = {
      Title: ORDocumentReview.Link,
      Request: ORDocumentReview.Request ? ORDocumentReview.Request : null,
      RequesttoId: ORDocumentReview.Requestto
        ? ORDocumentReview.Requestto
        : null,
      EmailccId: ORDocumentReview.Emailcc
        ? { results: ORDocumentReview.Emailcc }
        : { results: [] },
      Project: ORDocumentReview.Project ? ORDocumentReview.Project : null,
      Documenttype: ORDocumentReview.Documenttype
        ? ORDocumentReview.Documenttype
        : null,
      Comments: ORDocumentReview.Comments ? ORDocumentReview.Comments : null,
      Confidential: ORDocumentReview.Confidential,
      Product: ORDocumentReview.Product ? ORDocumentReview.Product : null,
      AnnualPlanID: ORDocumentReview.AnnualPlanID
        ? ORDocumentReview.AnnualPlanID
        : 0,
      DeliveryPlanID: ORDocumentReview.DeliveryPlanID
        ? ORDocumentReview.DeliveryPlanID
        : 0,
      ProductionBoardID: ORDocumentReview.ProductionBoardID
        ? ORDocumentReview.ProductionBoardID
        : null,
      DRPageName: "Org Reporting",
    };
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .items.add(requestdata)
      .then((e) => {
        if (ORDocumentReview.ProductionBoardID && !ORPendingPopup.condition) {
          sharepointWeb.lists
            .getByTitle("OrgReporting")
            .items.getById(ORDocumentReview.ProductionBoardID)
            .update({ Status: "Submitted" })
            .then(() => {
              if (ORPendingPopup.condition == false) {
                let Index = ORData.findIndex(
                  (obj) => obj.ID == ORDocumentReview.ProductionBoardID
                );
                ORData[Index].Status = "Submitted";
                setORData([...ORData]);
                sortORData = ORData;
              } else {
                let Index = ORPendingData.findIndex(
                  (obj) => obj.ID == ORDocumentReview.ProductionBoardID
                );
                ORPendingData[Index].Status = "Submitted";
                setORPendingData([...ORPendingData]);
                PendingData = ORPendingData;
              }
            })
            .catch(ORErrorFunction);
        }
        setORModalBoxVisibility(false);
        AddDRSuccessPopup();
      })
      .catch(ORErrorFunction);
  };
  const generateExcel = () => {
    let arrExport = ORFilterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Business area", key: "BA", width: 25 },
      { header: "Title", key: "Title", width: 25 },
      { header: "Frequency", key: "Frequency", width: 25 },
      { header: "Provider", key: "Provider", width: 25 },
      { header: "Audience", key: "Audience", width: 25 },
      { header: "Status", key: "Status", width: 60 },
      { header: "DueDate", key: "DueDate", width: 20 },
    ];
    arrExport.forEach((item) => {
      worksheet.addRow({
        BA: item.BA ? item.BA : "",
        Title: item.Title ? item.Title : "",
        Frequency: item.Frequency ? item.Frequency : "",
        Provider: item.Provider ? item.Provider : "",
        Audience: item.Audience ? item.Audience : "",
        Status: item.Status ? item.Status : "",
        DueDate: item.DueDate ? item.DueDate : "",
      });
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1"].map((key) => {
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
          `Organisationreport-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };
  const ORErrorFunction = (error: any, funName: string) => {
    console.log(error, funName);
    ErrorPopup();
  };
  const AddDRSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Document is successfully submitted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _ORColumns;
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

    const newORData = _copyAndSort(
      sortORData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newORFilterData = _copyAndSort(
      sortORFilterData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setORData([...newORData]);
    setORFilterData([...newORFilterData]);
    paginateFunction(1, [...newORFilterData]);
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
  const onChangeFilter = (key: string, option: string) => {
    let tempData: IData[] = ORData;
    let tempFilterKeys = ORFilter;
    tempFilterKeys[key] = option;
    if (tempFilterKeys.BA != "All") {
      tempData = tempData.filter((arr) => {
        return arr.BA == tempFilterKeys.BA;
      });
    }
    if (tempFilterKeys.Title != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Title == tempFilterKeys.Title;
      });
    }
    if (tempFilterKeys.Frequency != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Frequency == tempFilterKeys.Frequency;
      });
    }
    if (tempFilterKeys.Audience != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Audience == tempFilterKeys.Audience;
      });
    }
    if (tempFilterKeys.Status != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Status == tempFilterKeys.Status;
      });
    }
    setORFilterData([...tempData]);
    sortORFilterData = tempData;
    setORFilter({ ...tempFilterKeys });
    paginateFunction(1, tempData);
  };
  const GetUserDetails = (filterText) => {
    var result = allPeoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };
  const doesTextStartWith = (text: string, filterText: string) => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };
  const ORAddOnchange = (key, value) => {
    let tempArronchange = ORDocumentReview;
    if (key == "Request") tempArronchange.Request = value;
    else if (key == "Requestto") tempArronchange.Requestto = value;
    else if (key == "Emailcc") tempArronchange.Emailcc = value;
    else if (key == "Documenttype") tempArronchange.Documenttype = value;
    else if (key == "Link") tempArronchange.Link = value;
    else if (key == "Comments") tempArronchange.Comments = value;
    else if (key == "Confidential") tempArronchange.Confidential = value;

    setORDocumentReview(tempArronchange);
  };
  const paginateFunction = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      CurrentPage = pagenumber;
      setORDisplayData(paginatedItems);
      setORCurrentPage(pagenumber);
    } else {
      setORDisplayData([]);
      setORCurrentPage(1);
    }
  };
  // UseEffect
  useEffect(() => {
    setORLoader(true);
    getPendingReport();
    getModalBoxOptions();
  }, [ORReRender]);

  return (
    <div style={{ padding: "5px 15px" }}>
      {ORLoader ? <CustomLoader /> : null}
      <div
        style={{
          position: "sticky",
          top: 0,
          backgroundColor: "#fff",
          zIndex: 1,
          marginBottom: 10,
        }}
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
            <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
              Organization reporting
            </Label>
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "left",
            }}
          >
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
                marginRight: 10,
              }}
            >
              <Icon
                style={{
                  color: "#1D6F42",
                }}
                iconName="ExcelDocument"
                className={ORiconStyleClass.export}
              />
              Export as XLS
            </Label>
            {props.isAdmin ? (
              <a
                href={
                  `${props.URL}` + "/Lists/OrgReportingConfig/AllItems.aspx"
                }
                target="_blank"
              >
                <PrimaryButton
                  text="Add"
                  className={ORbuttonStyleClass.buttonPrimary}
                />
              </a>
            ) : null}
          </div>
        </div>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            marginBottom: 10,
          }}
        >
          <div className={styles.ddSection}>
            <div>
              <Label styles={ORLabelStyles}>Business area</Label>
              <Dropdown
                placeholder="Select an option"
                options={ORFilterDrpDown.BA}
                selectedKey={ORFilter.BA}
                styles={
                  ORFilter.BA == "All"
                    ? ORDropdownStyles
                    : ORActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("BA", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={ORLabelStyles}>Title</Label>
              <Dropdown
                placeholder="Select an option"
                options={ORFilterDrpDown.Title}
                selectedKey={ORFilter.Title}
                styles={
                  ORFilter.Title == "All"
                    ? ORDropdownStyles
                    : ORActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("Title", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={ORLabelStyles}>Frequency</Label>
              <Dropdown
                placeholder="Select an option"
                options={ORFilterDrpDown.Frequency}
                selectedKey={ORFilter.Frequency}
                styles={
                  ORFilter.Frequency == "All"
                    ? ORDropdownStyles
                    : ORActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("Frequency", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={ORLabelStyles}>Audience</Label>
              <Dropdown
                placeholder="Select an option"
                options={ORFilterDrpDown.Audience}
                selectedKey={ORFilter.Audience}
                styles={
                  ORFilter.Audience == "All"
                    ? ORDropdownStyles
                    : ORActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("Audience", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={ORLabelStyles}>Status</Label>
              <Dropdown
                placeholder="Select an option"
                options={ORFilterDrpDown.Status}
                selectedKey={ORFilter.Status}
                styles={
                  ORFilter.Status == "All"
                    ? ORDropdownStyles
                    : ORActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("Status", option["key"]);
                }}
              />
            </div>

            <div>
              <div>
                <Icon
                  iconName="Refresh"
                  title="Click to reset"
                  className={ORiconStyleClass.refresh}
                  onClick={() => {
                    setORFilterData([...ORMasterData]);
                    setORData([...ORMasterData]);
                    sortORFilterData = ORMasterData;
                    sortORData = ORMasterData;
                    paginateFunction(1, ORMasterData);
                    setORFilter({ ...ORFilterKeys });
                    setORColumns(_ORColumns);
                  }}
                />
              </div>
            </div>
          </div>
          <div
            style={{
              marginLeft: "20px",
              transform: "translateY(12px)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              margin: "0 10px",
            }}
          >
            <Label className={ORlabelStyles.NORLabel}>
              Number of records:{" "}
              <b style={{ color: "#038387" }}>{ORFilterData.length}</b>
            </Label>
          </div>
        </div>
      </div>
      <div>
        <DetailsList
          items={ORDisplayData}
          columns={ORColumns}
          // styles={mpDetailsListStyles}
          styles={{
            root: {
              ".ms-DetailsHeader-cellTitle": {
                // justifyContent: "center !important",
              },
              ".ms-DetailsRow-cell": {
                display: "flex",
                alignItems: "center",
                // justifyContent: "center",
              },
            },
          }}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
        />
      </div>
      {ORFilterData.length > 0 ? (
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            margin: "10px 0",
          }}
        >
          <Pagination
            currentPage={ORcurrentPage}
            totalPages={
              ORFilterData.length > 0
                ? Math.ceil(ORFilterData.length / totalPageItems)
                : 1
            }
            onChange={(page) => {
              paginateFunction(page, ORFilterData);
            }}
          />
        </div>
      ) : (
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            marginTop: "15px",
          }}
        >
          <Label style={{ color: "#2392B2" }}>No data Found !!!</Label>
        </div>
      )}
      <Modal isOpen={ORModalBoxVisibility} isBlocking={false}>
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
                errorMessage={ORShowMessage.Request}
                label="Request"
                placeholder="Select an option"
                options={ORModalBoxDropDownOptions.Request}
                styles={ORModalBoxDrpDwnCalloutStyles}
                onChange={(e, option: any) => {
                  ORAddOnchange("Request", option["key"]);
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
                className={ORModalBoxPP}
                onResolveSuggestions={GetUserDetails}
                itemLimit={1}
                onChange={(selectedUser) => {
                  selectedUser.length != 0
                    ? ORAddOnchange("Requestto", selectedUser[0]["ID"])
                    : ORAddOnchange("Requestto", "");
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
                {ORShowMessage.Requestto}
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
                className={ORModalBoxPP}
                onResolveSuggestions={GetUserDetails}
                itemLimit={5}
                onChange={(selectedUser) => {
                  let selectedId = selectedUser.map((su) => su["ID"]);
                  selectedUser.length != 0
                    ? ORAddOnchange("Emailcc", selectedId)
                    : ORAddOnchange("Emailcc", "");
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
                label="Title"
                // placeholder="Add new project"
                defaultValue={ORDocumentReview.Project}
                disabled={true}
                styles={ORTxtBoxStyles}
                className={styles.projectField}
                onChange={(e, value: string) => {}}
              />
            </div>
            <div>
              <Dropdown
                label="Document type"
                required={true}
                errorMessage={ORShowMessage.Documenttype}
                placeholder="Select an option"
                options={ORModalBoxDropDownOptions.Documenttype}
                styles={ORModalBoxDrpDwnCalloutStyles}
                onChange={(e, option: any) => {
                  ORAddOnchange("Documenttype", option["key"]);
                }}
              />
            </div>
            <div>
              <TextField
                label="Link"
                placeholder="Add link"
                errorMessage={ORShowMessage.Link}
                required={true}
                styles={ORTxtBoxStyles}
                onChange={(e, value: string) => {
                  ORAddOnchange("Link", value);
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
                label="Comments"
                placeholder="Add Comments"
                multiline
                rows={5}
                resizable={false}
                styles={ORMultiTxtBoxStyles}
                onChange={(e, value: string) => {
                  ORAddOnchange("Comments", value);
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
                  ORAddOnchange("Confidential", !ORDocumentReview.Confidential);
                }}
              />
            </div>
          </div>
          <div className={styles.apModalBoxButtonSection}>
            <button
              className={styles.apModalBoxSubmitBtn}
              onClick={(_) => {
                drValidationFunction();
              }}
              style={{ display: "flex" }}
            >
              {ORButtonLoader ? (
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
                setORModalBoxVisibility(false);
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
      {ORPendingPopup.condition ? (
        <Modal isOpen={ORPendingPopup.condition} isBlocking={false}>
          <div style={{ padding: "30px", paddingTop: "20px" }}>
            <Label className={styles.atpPopupLabel}>{"Pending reports"}</Label>
            <div
              style={{
                display: "flex",
                justifyContent: "flex-end",
              }}
            >
              <Icon
                iconName="Cancel"
                style={{
                  color: "#0882A5",
                  marginTop: -25,
                  marginRight: 10,
                  fontSize: 17,
                  height: 14,
                  width: 17,
                  cursor: "pointer",
                }}
                onClick={(_) => {
                  setORPendingPopup({
                    condition: false,
                    selectedItem: [],
                    pendingItem: [],
                  });
                }}
              />
            </div>

            <div>
              <div
                style={{
                  display: "flex",
                  justifyContent: "flex-start",
                  margin: "15px",
                  marginBottom: "10px",
                }}
              >
                <div style={{ width: "430px", paddingRight: "20px" }}>
                  <Label>Business area</Label>
                  {ORPendingPopup.selectedItem.length > 0
                    ? ORPendingPopup.selectedItem[0].BA
                    : null}
                </div>
                <div
                  style={{
                    width: "300px",
                    paddingRight: "20px",
                  }}
                >
                  <Label>Title</Label>
                  {ORPendingPopup.selectedItem.length > 0
                    ? ORPendingPopup.selectedItem[0].Title
                    : null}
                </div>
                <div style={{ width: "150px" }}>
                  <Label>Frequency</Label>
                  {ORPendingPopup.selectedItem.length > 0
                    ? ORPendingPopup.selectedItem[0].Frequency
                    : null}
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  justifyContent: "flex-start",
                  margin: "15px",
                  marginBottom: "10px",
                }}
              >
                {/* <div style={{ width: "430px", paddingRight: "20px" }}>
                  <Label>Responsible</Label>
                  {ORPendingPopup.selectedItem.length > 0
                    ? ORPendingPopup.selectedItem[0].Responsible
                    : null}
                </div> */}
                <div style={{ width: "430px", paddingRight: "20px" }}>
                  <Label>Provider</Label>
                  {ORPendingPopup.selectedItem.length > 0
                    ? ORPendingPopup.selectedItem[0].Provider
                    : null}
                </div>
                <div style={{ width: "300px" }}>
                  <Label>Audience</Label>
                  {ORPendingPopup.selectedItem.length > 0
                    ? ORPendingPopup.selectedItem[0].Audience
                    : null}
                </div>
              </div>
              <div
                style={{
                  marginTop: 30,
                  marginLeft: 15,
                  marginRight: 15,
                  width: 960,
                }}
              >
                {ORPendingPopup.pendingItem.length > 0 ? (
                  <DetailsList
                    items={ORPendingPopup.pendingItem}
                    columns={_pendingColumns}
                    styles={ORModalBoxDetailsListStyles}
                    setKey="set"
                    layoutMode={DetailsListLayoutMode.justified}
                    selectionMode={SelectionMode.none}
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
          </div>
        </Modal>
      ) : null}
    </div>
  );
};
export default OrgReporting;
