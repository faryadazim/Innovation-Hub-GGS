import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
import * as moment from "moment";
import {
  DetailsList,
  IDetailsListStyles,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  SearchBox,
  ISearchBoxStyles,
  TooltipHost,
  TooltipDelay,
  TooltipOverflowMode,
  DirectionalHint,
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
  Pivot,
  PivotItem,
} from "@fluentui/react";
import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
import { Log } from "@microsoft/sp-core-library";

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
interface IConfigData {
  ID: number;
  BA: string;
  Title: string;
  Frequency: string;
  Approver: string;
  Provider: string;
  Audience: string;
  TLink: string;
}
interface IData {
  ID: number;
  BA: string;
  Title: string;
  Frequency: string;
  ApproverDetails: any[];
  AudienceDetails: any[];
  Provider: any[];
  Status: string;
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
let ORPivot: number = 1;
const ORModalBoxDrpDwnOptns = {
  BA: [],
  Frequency: [],
};
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

const OrgReporting = (props: IProps): JSX.Element => {
  // Variable Declaration Starts
  const sharepointWeb: any = Web(props.URL);
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  const allPeoples: any[] = props.peopleList;
  let ORConfig = "Organisation reporting configuration list";

  const ORNewData = {
    ID: null,
    BA: null,
    Title: null,
    Frequency: null,
    ApproverDetails: [],
    AudienceDetails: [],
    Provider: "",
    Status: null,
    ConfigID: null,
    TimePeriod: null,
    Year: null,
    TLink: null,
    DueDate: null,

    BAValidation: false,
    TitleValidation: false,
    FrequencyValidation: false,
    ProviderValidation: false,
    AudienceValidation: false,
    ApproverValidation: false,

    overAllValidation: false,
  };
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
  const _ORAllReports: IColumn[] = [
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
      maxWidth: 100,
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
                cursor: "pointer",
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
    {
      key: "Column5",
      name: "Provider",
      fieldName: "Provider",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          {item.Provider.length > 0 ? (
            <>
              {
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "flex-start",
                    cursor: "pointer",
                  }}
                >
                  <div
                    title={item.Provider[0].text}
                    style={{ display: "flex" }}
                  >
                    <Persona
                      showOverflowTooltip
                      size={PersonaSize.size24}
                      presence={PersonaPresence.none}
                      showInitialsUntilImageLoads={true}
                      imageUrl={
                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                        `${item.Provider[0].secondaryText}`
                      }
                    />
                    {/* <Label style={{ marginLeft: 10 }}>
                      {item.Provider[0].text}
                    </Label> */}
                  </div>
                </div>
              }
            </>
          ) : (
            ""
          )}
        </>
      ),
    },
    {
      key: "Column6",
      name: "Audience",
      fieldName: "Audience",
      minWidth: 100,
      maxWidth: 200,
      onRender: (item) => (
        <>
          {item.AudienceDetails.length > 0 ? (
            <>
              {
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "flex-start",
                    cursor: "pointer",
                  }}
                >
                  <div title={item.AudienceDetails[0].text}>
                    <Persona
                      showOverflowTooltip
                      size={PersonaSize.size24}
                      presence={PersonaPresence.none}
                      showInitialsUntilImageLoads={true}
                      imageUrl={
                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                        `${item.AudienceDetails[0].secondaryText}`
                      }
                    />
                  </div>
                  {item.AudienceDetails.length > 1 ? (
                    <TooltipHost
                      content={
                        <ul style={{ margin: 10, padding: 0 }}>
                          {item.AudienceDetails.map((data, length) => {
                            if (length != 0) {
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
                                        `${data.secondaryText}`
                                      }
                                    />
                                    <Label style={{ marginLeft: 10 }}>
                                      {data.text}
                                    </Label>
                                  </div>
                                </li>
                              );
                            }
                          })}
                        </ul>
                      }
                      delay={TooltipDelay.zero}
                      id={item.ID}
                      directionalHint={DirectionalHint.bottomCenter}
                      styles={{ root: { display: "inline-block" } }}
                    >
                      <div
                        className={styles.extraPeople}
                        aria-describedby={item.ID}
                      >
                        {item.AudienceDetails.length - 1}
                      </div>
                    </TooltipHost>
                  ) : null}
                </div>
              }
            </>
          ) : (
            ""
          )}
        </>
      ),
    },
    {
      key: "Column7",
      name: "Approver",
      fieldName: "Approver",
      minWidth: 100,
      maxWidth: 200,
      onRender: (item) => (
        <>
          {item.ApproverDetails.length > 0 ? (
            <>
              {
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "flex-start",
                    cursor: "pointer",
                  }}
                >
                  <div title={item.ApproverDetails[0].text}>
                    <Persona
                      showOverflowTooltip
                      size={PersonaSize.size24}
                      presence={PersonaPresence.none}
                      showInitialsUntilImageLoads={true}
                      imageUrl={
                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                        `${item.ApproverDetails[0].secondaryText}`
                      }
                    />
                  </div>
                  {item.ApproverDetails.length > 1 ? (
                    <TooltipHost
                      content={
                        <ul style={{ margin: 10, padding: 0 }}>
                          {item.ApproverDetails.map((data, length) => {
                            if (length != 0) {
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
                                        `${data.secondaryText}`
                                      }
                                    />
                                    <Label style={{ marginLeft: 10 }}>
                                      {data.text}
                                    </Label>
                                  </div>
                                </li>
                              );
                            }
                          })}
                        </ul>
                      }
                      delay={TooltipDelay.zero}
                      id={item.ID}
                      directionalHint={DirectionalHint.bottomCenter}
                      styles={{ root: { display: "inline-block" } }}
                    >
                      <div
                        className={styles.extraPeople}
                        aria-describedby={item.ID}
                      >
                        {item.ApproverDetails.length - 1}
                      </div>
                    </TooltipHost>
                  ) : null}
                </div>
              }
            </>
          ) : (
            ""
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

      onRender: (item) => (
        <div style={{ display: "flex" }}>
          <div
            title="History"
            style={{
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              flexWrap: "wrap",
              width: 50,
            }}
          >
            <Icon
              iconName="Import"
              className={ORiconStyleClass.historyIcon}
              onClick={(): void => {
                console.log(item);
                setShowHistory({ condition: true, data: item });
              }}
            />
          </div>
        </div>
      ),
    },
  ];
  const _ORHistoryColumn: IColumn[] = [
    {
      key: "Column1",
      name: "Date",
      fieldName: "Date",
      minWidth: 120,
      maxWidth: 120,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column2",
      name: "Comments",
      fieldName: "Comments",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column3",
      name: "Document Link",
      fieldName: "Document Link",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column4",
      name: "Audience",
      fieldName: "Audience",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column5",
      name: "Action",
      fieldName: "Action",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
  ];
  const _ORMyReportsColumn: IColumn[] = [
    {
      key: "Column1",
      name: "BA",
      fieldName: "BA",
      minWidth: 120,
      maxWidth: 120,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
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
      name: "Due Date",
      fieldName: "Due Date",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column5",
      name: "Provider",
      fieldName: "Provider",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column6",
      name: "Audience",
      fieldName: "Audience",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column7",
      name: "Status",
      fieldName: "Status",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column8",
      name: "TL",
      fieldName: "TL",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column9",
      name: "File Upload",
      fieldName: "File Upload",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
  ];
  // Variable Declaration Ends
  // Style Section Starts
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
    refresh: {
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
    pblink: {
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
    export: {
      color: "black",
      fontSize: "18px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
    },
    historyIcon: {
      color: "#000",
      fontSize: "18px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      fontWeight: 600,
    },
    historyBackIcon: {
      color: "#000",
      fontSize: "18px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      fontWeight: 600,
      marginTop: 8,
    },
  });
  const ORlabelStyles = mergeStyleSets({
    titleLabel: {
      color: "#676767",
      fontSize: "14px",
      marginRight: "10px",
      fontWeight: "400",
    },
    selectedLabel: {
      color: "#0882A5",
      fontSize: "14px",
      marginRight: "10px",
      fontWeight: "600",
    },
    labelValue: {
      color: "#0882A5",
      fontSize: "14px",
      marginRight: "10px",
    },
    inputLabels: {
      color: "#323130",
      fontSize: "13px",
    },
    ErrorLabel: {
      marginTop: "25px",
      marginLeft: "10px",
      fontWeight: "500",
      color: "#D0342C",
      fontSize: "13px",
    },
    NORLabel: {
      color: "#323130",
      fontSize: "13px",
      marginLeft: "10px",
      fontWeight: "500",
    },
    historyHeading: {
      color: "#323130",
      fontSize: 16,
      marginLeft: 10,
      fontWeight: 600,
    },
    historyDescHeading1: { color: "#000", fontSize: 16, fontWeight: 500 },
    historyDescHeading2: {
      color: "#2392B2",
      fontSize: 16,
      fontWeight: 500,
      marginLeft: 5,
    },
  });
  const ORLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const ORModalBoxDropDownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      margin: "10px 20px",
      backgroundColor: "#fff",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      borderRadius: 4,
      border: "1px solid #000",
      color: "#000",
    },
    dropdownItemSelected: { fontSize: 12, backgroundColor: "#fff" },
    caretDown: {
      fontSize: 14,
      color: "#000",
    },
    callout: { height: 200 },
  };
  const ORModalBoxErrorDropDownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      margin: "10px 20px",
      backgroundColor: "#fff",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      borderRadius: 4,
      border: "2px solid #f00",
      color: "#000",
    },
    dropdownItemSelected: { fontSize: 12, backgroundColor: "#fff" },
    caretDown: {
      fontSize: 14,
      color: "#000",
    },
    callout: { height: 200 },
  };
  const ORModalBoxPP = mergeStyles({
    width: "300px",
    margin: "10px 20px",
  });
  const ORModalBoxTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "300px",
      margin: "10px 20px",
    },
    field: {
      fontSize: 12,
      color: "#000",
    },
    fieldGroup: {
      border: "1px solid #000",
    },
  };
  const ORModalBoxErrorTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "300px",
      margin: "10px 20px",
    },
    field: {
      fontSize: 12,
      color: "#000",
    },
    fieldGroup: {
      border: "2px solid #f00",
    },
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
  // Style Section Ends
  // State Declaration Starts
  const [ORReRender, setORReRender] = useState<boolean>(false);
  const [ORMasterData, setORMasterData] = useState<IData[]>([]);
  const [ORData, setORData] = useState<IData[]>([]);
  const [ORFilterData, setORFilterData] = useState<IData[]>([]);
  const [ORDisplayData, setORDisplayData] = useState<IData[]>([]);
  const [ORLoader, setORLoader] = useState<boolean>(false);
  const [showHistory, setShowHistory] = useState<{
    condition: boolean;
    data: IData;
  }>({ condition: false, data: null });
  const [ORFilter, setORFilter] = useState<IFilter>(ORFilterKeys);
  const [ORFilterDrpDown, setORFilterDrpDown] =
    useState<IDropdowns>(ORFilterOptns);
  const [ORAddConfigModalBox, setORAddConfigModalBox] = useState({
    visible: false,
    value: ORNewData,
  });
  const [ORButtonLoader, setORButtonLoader] = useState<boolean>(false);
  const [ORColumns, setORColumns] = useState([]);
  const [ORcurrentPage, setORCurrentPage] = useState<number>(CurrentPage);
  const [ORModalBoxDropDownOptions, setORModalBoxDropDownOptions] = useState(
    ORModalBoxDrpDwnOptns
  );
  // State Declaration Ends

  // Function Declaration Starts
  // common function
  const ORValidationFunction = () => {
    let tempORAddConfigModalBox = { ...ORAddConfigModalBox };
    if (!tempORAddConfigModalBox.value.BA) {
      tempORAddConfigModalBox.value.BAValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }
    if (!tempORAddConfigModalBox.value.Title) {
      tempORAddConfigModalBox.value.TitleValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }
    if (!tempORAddConfigModalBox.value.Frequency) {
      tempORAddConfigModalBox.value.FrequencyValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }
    if (tempORAddConfigModalBox.value.ApproverDetails.length <= 0) {
      tempORAddConfigModalBox.value.ApproverValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }
    if (tempORAddConfigModalBox.value.AudienceDetails.length <= 0) {
      tempORAddConfigModalBox.value.AudienceValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }
    if (!tempORAddConfigModalBox.value.Provider) {
      tempORAddConfigModalBox.value.ProviderValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }

    if (tempORAddConfigModalBox.value.overAllValidation) {
      setORButtonLoader(false);
      setORAddConfigModalBox({ ...tempORAddConfigModalBox });
    } else {
      let _ApproverDetails = [];
      let _AudienceDetails = [];
      if (ORAddConfigModalBox.value.ApproverDetails.length > 0) {
        ORAddConfigModalBox.value.ApproverDetails.forEach((_data) => {
          _ApproverDetails.push(_data.ID);
        });
      }
      if (ORAddConfigModalBox.value.AudienceDetails.length > 0) {
        ORAddConfigModalBox.value.AudienceDetails.forEach((_data) => {
          _AudienceDetails.push(_data.ID);
        });
      }
      let responseData = {
        BA: ORAddConfigModalBox.value.BA,
        Title: ORAddConfigModalBox.value.Title,
        Frequency: ORAddConfigModalBox.value.Frequency,
        ApproverId:
          ORAddConfigModalBox.value.ApproverDetails.length > 0
            ? { results: [..._ApproverDetails] }
            : { results: [] },
        AudienceId:
          ORAddConfigModalBox.value.AudienceDetails.length > 0
            ? { results: [..._AudienceDetails] }
            : { results: [] },
        ProviderId: ORAddConfigModalBox.value.Provider
          ? ORAddConfigModalBox.value.Provider
          : null,
      };
      console.log(responseData);
      ORAddFunction("addConfig", responseData);
    }
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

  // sorting function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempORColumns =
      ORPivot == 1 ? _ORAllReports : ORPivot == 2 ? _ORMyReportsColumn : null;
    const newColumns: IColumn[] = tempORColumns.slice();
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

  // getData functions
  const getOrgReportConfig = (): void => {
    setORLoader(true);
    setORColumns(_ORAllReports);
    let _ORdata: IData[] = [];
    sharepointWeb.lists
      .getByTitle(ORConfig)
      .items.top(5000)
      .orderBy("Modified", false)
      .get()
      .then(async (items) => {
        items.forEach((item) => {
          let _ApproverDetails = [];
          let _AudienceDetails = [];
          let _ProviderDetails = [];
          if (item.ApproverId.length > 0) {
            item.ApproverId.forEach((user) => {
              _ApproverDetails.push(
                allPeoples.filter((people) => {
                  return people.ID == user;
                })[0]
              );
            });
          }
          if (item.AudienceId.length > 0) {
            item.AudienceId.forEach((user) => {
              _AudienceDetails.push(
                allPeoples.filter((people) => {
                  return people.ID == user;
                })[0]
              );
            });
          }
          if (item.ProviderId) {
            _ProviderDetails.push(
              allPeoples.filter((people) => {
                return people.ID == item.ProviderId;
              })[0]
            );
          }
          _ORdata.push({
            ID: item.ID,
            BA: item.BA,
            Title: item.Title,
            Frequency: item.Frequency,
            ApproverDetails: [..._ApproverDetails],
            AudienceDetails: [..._AudienceDetails],
            Provider: [..._ProviderDetails],
            Status: item.Status,
            ConfigID: item.ConfigID,
            TimePeriod: item.TimePeriod,
            Year: item.Year,
            TLink: item.TLink,
            DueDate: item.DueDate,
          });
        });
        console.log(_ORdata);

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
        ORErrorFunction(error, "getOrgReportConfig");
      });
  };
  const getOrgReporting = (): void => {
    console.log("set7");
    setORLoader(true);
    console.log("set8");
    setORColumns(_ORMyReportsColumn);
    let _ORdata: IData[] = [];
    sharepointWeb.lists
      .getByTitle("OrgReporting")
      .items.top(5000)
      .orderBy("Modified", false)
      .get()
      .then(async (items) => {
        console.log(items.length);

        items.forEach((item) => {
          let _ApproverDetails = [];
          let _AudienceDetails = [];
          if (item.ApproverId.length > 0) {
            item.Approver.forEach((user) => {
              _ApproverDetails.push(
                allPeoples.filter((people) => {
                  return people.ID == user.Id;
                })[0]
              );
            });
          }
          if (item.AudienceId.length > 0) {
            item.Audience.forEach((user) => {
              _AudienceDetails.push(
                allPeoples.filter((people) => {
                  return people.ID == user.Id;
                })[0]
              );
            });
          }
          _ORdata.push({
            ID: item.ID,
            BA: item.BA,
            Title: item.Title,
            Frequency: item.Frequency,
            ApproverDetails: [..._ApproverDetails],
            AudienceDetails: [..._AudienceDetails],
            Provider: item.Provider ? item.Provider.Title : null,
            Status: item.Status,
            ConfigID: item.ConfigID,
            TimePeriod: item.TimePeriod,
            Year: item.Year,
            TLink: item.TLink,
            DueDate: item.DueDate,
          });
        });
        console.log("set1");
        setORFilterData([..._ORdata]);
        console.log("set2");
        sortORFilterData = _ORdata;
        setORData([..._ORdata]);
        sortORData = _ORdata;
        console.log("set3");
        setORMasterData([..._ORdata]);
        console.log("set4");
        reloadFilterDropdowns([..._ORdata]);
        console.log("set5");
        paginateFunction(1, _ORdata);
        console.log("set6");
        setORLoader(false);
      })
      .catch((error) => {
        ORErrorFunction(error, "getOrgReporting");
      });
  };

  const getModalBoxOptions = () => {
    //Request Choices
    sharepointWeb.lists
      .getByTitle(ORConfig)
      .fields.getByInternalNameOrTitle("BA")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              ORModalBoxDrpDwnOptns.BA.findIndex((rpb) => {
                return rpb.key == choice;
              }) == -1
            ) {
              ORModalBoxDrpDwnOptns.BA.push({
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
      .getByTitle(ORConfig)
      .fields.getByInternalNameOrTitle("Frequency")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              ORModalBoxDrpDwnOptns.Frequency.findIndex((rdt) => {
                return rdt.key == choice;
              }) == -1
            ) {
              ORModalBoxDrpDwnOptns.Frequency.push({
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
  const ORAddOnchange = (key, value) => {
    let tempArronchange = { ...ORAddConfigModalBox.value };
    tempArronchange.BAValidation = false;
    if (key == "BA") {
      tempArronchange.BA = value;
      tempArronchange.BAValidation = false;
    } else if (key == "Title") {
      tempArronchange.Title = value;
      tempArronchange.TitleValidation = false;
    } else if (key == "Frequency") {
      tempArronchange.Frequency = value;
      tempArronchange.FrequencyValidation = false;
    } else if (key == "ApproverDetails") {
      tempArronchange.ApproverDetails = value;
      tempArronchange.ApproverValidation = false;
    } else if (key == "Provider") {
      tempArronchange.Provider = value;
      tempArronchange.ProviderValidation = false;
    } else if (key == "AudienceDetails") {
      tempArronchange.AudienceDetails = value;
      tempArronchange.AudienceValidation = false;
    }

    setORAddConfigModalBox({
      visible: true,
      value: tempArronchange,
    });
  };

  // Onchange Function
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
      let devArr = [];
      tempData.forEach((arr) => {
        if (arr.AudienceDetails.length != 0) {
          if (
            arr.AudienceDetails.some(
              (people) => people.text == tempFilterKeys.Audience
            )
          ) {
            devArr.push(arr);
          }
        }
      });
      tempData = [...devArr];
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
  const reloadFilterDropdowns = (data: IData[]): void => {
    console.log(data);
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
        ORFilterOptns.BA.push({
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
        ORFilterOptns.Title.push({
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
        ORFilterOptns.Frequency.push({
          key: item.Frequency,
          text: item.Frequency,
        });
      }
      let tempAudience = [];
      if (item.AudienceDetails.length > 0) {
        item.AudienceDetails.forEach((people) => {
          tempAudience.push(people.text);
        });

        tempAudience.forEach((_people) => {
          if (
            ORFilterOptns.Audience.findIndex((audienceOptns) => {
              return audienceOptns.key == _people;
            }) == -1 &&
            _people != null
          ) {
            ORFilterOptns.Audience.push({
              key: _people,
              text: _people,
            });
          }
        });
      }
      if (
        ORFilterOptns.Status.findIndex((Status) => {
          return Status.key == item.Status;
        }) == -1 &&
        item.Status
      ) {
        ORFilterOptns.Status.push({
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

  const ORAddFunction = (key, responseData) => {
    sharepointWeb.lists
      .getByTitle("Organisation reporting configuration list")
      .items.add(responseData)
      .then(() => {
        if (key == "addConfig") {
          AddDRSuccessPopup();
          setORButtonLoader(false);
          setORAddConfigModalBox({
            visible: false,
            value: ORNewData,
          });
          setORReRender(!ORReRender);
        }
      })
      .catch(ORErrorFunction);
  };

  // return function
  const PivotFunction = (param1, param2): JSX.Element => {
    console.log("test");
    return (
      <div style={{ marginTop: 15, padding: 10 }}>
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
                      setORColumns(_ORAllReports);
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
            items={param1}
            columns={param2}
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
      </div>
    );
  };

  const historyPage = (): JSX.Element => {
    return (
      <div style={{ border: "1px solid black", padding: 10, marginTop: 15 }}>
        <div style={{ display: "flex" }}>
          <Icon
            iconName="ChromeBack"
            className={ORiconStyleClass.historyBackIcon}
            onClick={(): void => {
              setShowHistory({ condition: false, data: null });
            }}
          />
          <Label className={ORlabelStyles.historyHeading}>History</Label>
        </div>
        <div style={{ display: "flex" }}>
          <div style={{ display: "flex", marginRight: 15 }}>
            <Label className={ORlabelStyles.historyDescHeading1}>
              Business area :
            </Label>
            <Label className={ORlabelStyles.historyDescHeading2}>
              {showHistory.data.BA}
            </Label>
          </div>
          <div style={{ display: "flex", marginRight: 15 }}>
            <Label className={ORlabelStyles.historyDescHeading1}>Title :</Label>
            <Label className={ORlabelStyles.historyDescHeading2}>
              {showHistory.data.Title}
            </Label>
          </div>
          <div style={{ display: "flex", marginRight: 15 }}>
            <Label className={ORlabelStyles.historyDescHeading1}>
              Frequency :
            </Label>
            <Label className={ORlabelStyles.historyDescHeading2}>
              {showHistory.data.Frequency}
            </Label>
          </div>
        </div>
        <div>
          <DetailsList
            items={ORDisplayData}
            columns={ORColumns}
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
      </div>
    );
  };
  // Function Declaration Ends
  useEffect(() => {
    getModalBoxOptions();
    ORPivot == 1
      ? getOrgReportConfig()
      : ORPivot == 2
      ? getOrgReporting()
      : ORPivot == 3
      ? getOrgReporting()
      : null;
  }, [ORReRender]);

  console.log("Hi return");

  return (
    <>
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
                Organisation reporting
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
              {props.isAdmin && ORPivot == 1 ? (
                <PrimaryButton
                  text="Add"
                  className={ORbuttonStyleClass.buttonPrimary}
                  onClick={(_) =>
                    setORAddConfigModalBox({
                      visible: true,
                      value: ORNewData,
                    })
                  }
                />
              ) : null}
            </div>
          </div>
        </div>
        <Pivot
          onLinkClick={(e: any) => {
            console.log(e.props.headerText);
            if (e.props.headerText == "All reports") {
              ORPivot = 1;
              setShowHistory({ condition: false, data: null });
              setORReRender(!ORReRender);
            } else if (e.props.headerText == "My reports") {
              ORPivot = 2;
              setShowHistory({ condition: false, data: null });
              setORReRender(!ORReRender);
            } else if (e.props.headerText == "Approval requests") {
              ORPivot = 3;
              setShowHistory({ condition: false, data: null });
              setORReRender(!ORReRender);
            }
          }}
        >
          <PivotItem headerText="All reports">
            {showHistory.condition
              ? historyPage()
              : PivotFunction([], _ORAllReports)}
          </PivotItem>
          <PivotItem headerText="My reports">
            {PivotFunction([], _ORMyReportsColumn)}
          </PivotItem>
          <PivotItem headerText="Approval requests">
            {PivotFunction([], [])}
          </PivotItem>
        </Pivot>

        {ORAddConfigModalBox.visible ? (
          <Modal isOpen={ORAddConfigModalBox.visible} isBlocking={false}>
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
                Add organisation reporting
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
                    required={true}
                    placeholder="Select an option"
                    options={ORModalBoxDropDownOptions.BA}
                    styles={
                      ORAddConfigModalBox.value.BAValidation
                        ? ORModalBoxErrorDropDownStyles
                        : ORModalBoxDropDownStyles
                    }
                    selectedKey={ORAddConfigModalBox.value.BA}
                    onChange={(e, option: any) => {
                      ORAddOnchange("BA", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <TextField
                    label="Title"
                    placeholder="Add new project"
                    value={ORAddConfigModalBox.value.Title}
                    styles={
                      ORAddConfigModalBox.value.BAValidation
                        ? ORModalBoxErrorTxtBoxStyles
                        : ORModalBoxTxtBoxStyles
                    }
                    className={styles.projectField}
                    onChange={(e, value: string) => {
                      ORAddOnchange("Title", value);
                    }}
                  />
                </div>
                <div>
                  <Dropdown
                    required={true}
                    label="Frequency"
                    placeholder="Select an option"
                    options={ORModalBoxDropDownOptions.Frequency}
                    styles={
                      ORAddConfigModalBox.value.FrequencyValidation
                        ? ORModalBoxErrorDropDownStyles
                        : ORModalBoxDropDownStyles
                    }
                    selectedKey={ORAddConfigModalBox.value.Frequency}
                    onChange={(e, option: any) => {
                      ORAddOnchange("Frequency", option["key"]);
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
                  <Label
                    required={true}
                    style={{
                      transform: "translate(20px, 10px)",
                    }}
                  >
                    Provider
                  </Label>
                  <NormalPeoplePicker
                    className={ORModalBoxPP}
                    styles={
                      ORAddConfigModalBox.value.ProviderValidation
                        ? {
                            root: {
                              selectors: {
                                ".ms-BasePicker-text": {
                                  border: "2px solid #f00",
                                },
                              },
                            },
                          }
                        : {
                            root: {
                              selectors: {
                                selectors: {
                                  ".ms-BasePicker-text": {
                                    border: "1px solid #000",
                                  },
                                },
                              },
                            },
                          }
                    }
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={1}
                    selectedItems={allPeoples.filter((people) => {
                      return people.ID == ORAddConfigModalBox.value.Provider;
                    })}
                    onChange={(selectedUser) => {
                      selectedUser.length != 0
                        ? ORAddOnchange("Provider", selectedUser[0]["ID"])
                        : ORAddOnchange("Provider", "");
                    }}
                  />
                </div>
                <div>
                  <Label
                    style={{
                      transform: "translate(20px, 10px)",
                    }}
                  >
                    Audience
                  </Label>
                  <NormalPeoplePicker
                    className={ORModalBoxPP}
                    styles={
                      ORAddConfigModalBox.value.AudienceValidation
                        ? {
                            root: {
                              selectors: {
                                ".ms-BasePicker-text": {
                                  border: "2px solid #f00",
                                },
                              },
                            },
                          }
                        : {
                            root: {
                              selectors: {
                                selectors: {
                                  ".ms-BasePicker-text": {
                                    border: "1px solid #000",
                                  },
                                },
                              },
                            },
                          }
                    }
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={5}
                    selectedItems={ORAddConfigModalBox.value.AudienceDetails}
                    onChange={(selectedUser) => {
                      selectedUser.length != 0
                        ? ORAddOnchange("AudienceDetails", selectedUser)
                        : ORAddOnchange("AudienceDetails", "");
                    }}
                  />
                </div>
                <div>
                  <Label
                    style={{
                      transform: "translate(20px, 10px)",
                    }}
                  >
                    Approver
                  </Label>
                  <NormalPeoplePicker
                    className={ORModalBoxPP}
                    styles={
                      ORAddConfigModalBox.value.ApproverValidation
                        ? {
                            root: {
                              selectors: {
                                ".ms-BasePicker-text": {
                                  border: "2px solid #f00",
                                },
                              },
                            },
                          }
                        : {
                            root: {
                              selectors: {
                                selectors: {
                                  ".ms-BasePicker-text": {
                                    border: "1px solid #000",
                                  },
                                },
                              },
                            },
                          }
                    }
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={5}
                    selectedItems={ORAddConfigModalBox.value.ApproverDetails}
                    onChange={(selectedUser) => {
                      selectedUser.length != 0
                        ? ORAddOnchange("ApproverDetails", selectedUser)
                        : ORAddOnchange("ApproverDetails", "");
                    }}
                  />
                </div>
              </div>
              <div className={styles.apModalBoxButtonSection}>
                {ORAddConfigModalBox.value.overAllValidation ? (
                  <Label style={{ color: "#f00", fontWeight: 600 }}>
                    * All fields are mandatory
                  </Label>
                ) : null}
                <button
                  className={styles.apModalBoxSubmitBtn}
                  onClick={(_) => {
                    setORButtonLoader(true);
                    ORValidationFunction();
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
                    setORAddConfigModalBox({
                      visible: false,
                      value: ORNewData,
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
        ) : null}
      </div>
    </>
  );
};
export default OrgReporting;
