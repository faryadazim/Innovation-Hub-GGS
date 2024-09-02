import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
import * as moment from "moment";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IDetailsColumnStyles,
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
  IStackTokens,
  Stack,
  Toggle,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./WeeklyReport.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
import CustomLoader from "./CustomLoader";
import Pagination from "office-ui-fabric-react-pagination";
import PHDashboardTable from "./DashboardTable";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

interface IProps {
  context: WebPartContext;
  spcontext: any;
  graphContent: any;
  URL: string;
  peopleList: any;
  BA: string;
}
interface IFilter {
  User: string;
  Week: string;
  Year: string;
  UserState: string;
  BA: string;
  ShowAll: boolean;
}
interface IDropdown {
  key: string;
  text: string;
}
interface IDropdownOptions {
  User: IDropdown[];
  Week: IDropdown[];
  Year: IDropdown[];
  UserState: IDropdown[];
  BA: IDropdown[];
}
interface projectDetails {
  ID: number;
  PBType: string;
  projectName: string;
  AH: number;
  PH: number;
  EndDate: string;
}
interface IData {
  ID: number;

  UserID: number;
  UserName: string;
  UserEmail: string;

  maxProjects: projectDetails[];

  BA: string;
  ActiveStatus: string;

  TH: number;
  AH: number;
  PH: number;

  ShowAll: boolean;

  Review: number;
  Edit: number;
  Assemble: number;
  SignOff: number;
  Publish: number;
  Returned: number;
  Feedback: number;
  Actioned: number;
  Endosed: number;
  SignedOff: number;
  RepeatedIssues: number;
  Major_Moderate: number;
  Incomplete: number;
  Quality: number;
  Style: number;

  ReviewData: IHistoryData[];
  EditData: IHistoryData[];
  AssembleData: IHistoryData[];
  SignOffData: IHistoryData[];
  PublishData: IHistoryData[];
  ReturnedData: IHistoryData[];
  FeedbackData: IHistoryData[];
  ActionedData: IHistoryData[];
  EndosedData: IHistoryData[];
  SignedOffData: IHistoryData[];
  RepeatedIssuesData: IHistoryData[];
  Major_ModerateData: IHistoryData[];
  IncompleteData: IHistoryData[];
  QualityData: IHistoryData[];
  StyleData: IHistoryData[];
}
interface IHistoryData {
  PBID: number;
  PBType: string;

  FileLink: string;
  FileName: string;
  Sent: string;
  SentToName: string;

  FromName: string;
  ResponseDate: Date;
  Rating: number;
  Requests: string;
  Responses: string;
  ResponseComments: string;
  RequestComments: string;
  From: { ID: number; Name: string; Email: string };
  SentTo: { ID: number; Name: string; Email: string };
}

let sortData: IData[] = [];
let sortFilterData: IData[] = [];

let sortHistoryData: IHistoryData[] = [];

let globalMasterUserListData = [];
let globalPBData = [];
let globalAPBData = [];
let globalDRData = [];

let CurrentPage: number = 1;
let totalPageItems: number = 10;

const WRDashboard = (props: IProps) => {
  // variable-Declaration Starts
  const sharepointWeb: any = Web(props.URL);
  const allPeoples: any[] = props.peopleList;
  const currentBA = props.BA;

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  const currentYear: number = moment().year();
  const currentWeekNumber: number = moment().isoWeek();

  const docReviewHeaderStyle: Partial<IDetailsColumnStyles> = {
    cellName: {
      color: "#038387",
    },
  };
  const docResponseHeaderStyle: Partial<IDetailsColumnStyles> = {
    cellName: {
      color: "#FAA332",
    },
  };
  const docQualityFeedbackHeaderStyle: Partial<IDetailsColumnStyles> = {
    cellName: {
      color: "#FA6232",
    },
  };

  const _DBColumns: any[] = [
    {
      key: "Column1",
      name: "User",
      fieldName: "UserName",
      minWidth: 200,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div style={{ display: "flex" }}>
          <div style={{ cursor: "pointer" }}>
            <Persona
              title={item.UserName}
              size={PersonaSize.size24}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.UserEmail}`
              }
            />
          </div>
          <Label style={{ marginTop: -3 }}>{item.UserName}</Label>
        </div>
      ),
    },
    {
      key: "Column3",
      name: (
        <TooltipHost content="Actual hours">
          <span style={{ cursor: "pointer" }}>AH</span>
        </TooltipHost>
      ),
      fieldName: "AH",
      minWidth: 40,
      maxWidth: 60,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div>
          {item.AH.toString().match(/\./g) ? item.AH.toFixed(2) : item.AH}
        </div>
      ),
    },
    {
      key: "Column4",
      name: (
        <TooltipHost content="Planned hours">
          <span style={{ cursor: "pointer" }}>PH</span>
        </TooltipHost>
      ),
      fieldName: "PH",
      minWidth: 40,
      maxWidth: 60,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div>
          {item.PH.toString().match(/\./g) ? item.PH.toFixed(2) : item.PH}
        </div>
      ),
    },
    {
      key: "Column7",
      name: "Review",
      fieldName: "Review",
      minWidth: 50,
      maxWidth: 70,

      styles: docReviewHeaderStyle,

      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{ cursor: "pointer" }}
          onClick={() => {
            setDBHistoryColumns(_DBHistoryColumns);
            setDetailHistory({
              condition: true,
              userName: item.UserName,
              userEmail: item.UserEmail,
              type: "Review",
              data: item.ReviewData,
            });
            setFilteredHistoryData([...item.ReviewData]);
            sortHistoryData = [...item.ReviewData];
          }}
        >{`${item.Review}`}</div>
      ),
    },
    {
      key: "Column8",
      name: "Edit",
      fieldName: "Edit",
      minWidth: 50,
      maxWidth: 60,

      styles: docReviewHeaderStyle,

      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{ cursor: "pointer" }}
          onClick={() => {
            setDBHistoryColumns(_DBHistoryColumns);
            setDetailHistory({
              condition: true,
              userName: item.UserName,
              userEmail: item.UserEmail,
              type: "Edit",
              data: item.EditData,
            });
            setFilteredHistoryData([...item.EditData]);
            sortHistoryData = [...item.EditData];
          }}
        >{`${item.Edit}`}</div>
      ),
    },
    {
      key: "Column9",
      name: "Assemble",
      fieldName: "Assemble",
      minWidth: 80,
      maxWidth: 100,

      styles: docReviewHeaderStyle,

      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{ cursor: "pointer" }}
          onClick={() => {
            setDBHistoryColumns(_DBHistoryColumns);
            setDetailHistory({
              condition: true,
              userName: item.UserName,
              userEmail: item.UserEmail,
              type: "Assemble",
              data: item.AssembleData,
            });
            setFilteredHistoryData([...item.AssembleData]);
            sortHistoryData = [...item.AssembleData];
          }}
        >{`${item.Assemble}`}</div>
      ),
    },
    {
      key: "Column10",
      name: "Sign off",
      fieldName: "SignOff",
      minWidth: 60,
      maxWidth: 100,

      styles: docReviewHeaderStyle,

      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{ cursor: "pointer" }}
          onClick={() => {
            setDBHistoryColumns(_DBHistoryColumns);
            setDetailHistory({
              condition: true,
              userName: item.UserName,
              userEmail: item.UserEmail,
              type: "Sign off",
              data: item.SignOffData,
            });
            setFilteredHistoryData([...item.SignOffData]);
            sortHistoryData = [...item.SignOffData];
          }}
        >{`${item.SignOff}`}</div>
      ),
    },
    {
      key: "Column11",
      name: "Returned",
      fieldName: "Returned",
      minWidth: 70,
      maxWidth: 100,

      styles: docResponseHeaderStyle,

      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{ cursor: "pointer" }}
          onClick={() => {
            setDBHistoryColumns(_DBHistoryColumns);
            setDetailHistory({
              condition: true,
              userName: item.UserName,
              userEmail: item.UserEmail,
              type: "Returned",
              data: item.ReturnedData,
            });
            setFilteredHistoryData([...item.ReturnedData]);
            sortHistoryData = [...item.ReturnedData];
          }}
        >{`${item.Returned}`}</div>
      ),
    },
    {
      key: "Column12",
      name: "Feedback",
      fieldName: "Feedback",
      minWidth: 70,
      maxWidth: 100,

      styles: docResponseHeaderStyle,

      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{ cursor: "pointer" }}
          onClick={() => {
            setDBHistoryColumns(_DBHistoryColumns);
            setDetailHistory({
              condition: true,
              userName: item.UserName,
              userEmail: item.UserEmail,
              type: "Feedback",
              data: item.FeedbackData,
            });
            setFilteredHistoryData([...item.FeedbackData]);
            sortHistoryData = [...item.FeedbackData];
          }}
        >{`${item.Feedback}`}</div>
      ),
    },
    {
      key: "Column14",
      name: "Endosed",
      fieldName: "Endosed",
      minWidth: 60,
      maxWidth: 100,

      styles: docResponseHeaderStyle,

      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{ cursor: "pointer" }}
          onClick={() => {
            setDBHistoryColumns(_DBHistoryColumns);
            setDetailHistory({
              condition: true,
              userName: item.UserName,
              userEmail: item.UserEmail,
              type: "Endosed",
              data: item.EndosedData,
            });
            setFilteredHistoryData([...item.EndosedData]);
            sortHistoryData = [...item.EndosedData];
          }}
        >{`${item.Endosed}`}</div>
      ),
    },
    {
      key: "Column15",
      name: "Signed off",
      fieldName: "SignedOff",
      minWidth: 70,
      maxWidth: 100,

      styles: docResponseHeaderStyle,

      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{ cursor: "pointer" }}
          onClick={() => {
            setDBHistoryColumns(_DBHistoryColumns);
            setDetailHistory({
              condition: true,
              userName: item.UserName,
              userEmail: item.UserEmail,
              type: "Signed off",
              data: item.SignedOffData,
            });
            setFilteredHistoryData([...item.SignedOffData]);
            sortHistoryData = [...item.SignedOffData];
          }}
        >{`${item.SignedOff}`}</div>
      ),
    },
    {
      key: "Column17",
      name: "Repeated issues",
      fieldName: "RepeatedIssues",
      minWidth: 120,
      maxWidth: 250,

      styles: docQualityFeedbackHeaderStyle,

      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{ cursor: "pointer" }}
          onClick={() => {
            setDBHistoryColumns(_DBHistoryColumns);
            setDetailHistory({
              condition: true,
              userName: item.UserName,
              userEmail: item.UserEmail,
              type: "Repeated Issues",
              data: item.RepeatedIssuesData,
            });
            setFilteredHistoryData([...item.RepeatedIssuesData]);
            sortHistoryData = [...item.RepeatedIssuesData];
          }}
        >{`${item.RepeatedIssues}`}</div>
      ),
    },
  ];

  const _DBHistoryColumns: IColumn[] = [
    {
      key: "Column1",
      name: "File Submitted",
      fieldName: "FromName",
      minWidth: 75,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onHistoryColumnClick(ev, column);
      },
      onRender: (item) => <div title={item.FromName}>{item.FromName}</div>,
    },
    {
      key: "Column2",
      name: "Actual sent date",
      fieldName: "Sent",
      minWidth: 60,
      maxWidth: 120,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onHistoryColumnClick(ev, column);
      },
      onRender: (item) => <div>{moment(item.Sent).format("DD/MM/YYYY")}</div>,
    },
    {
      key: "Column3",
      name: "Title",
      fieldName: "FileName",
      isResizable: true,
      minWidth: 125,
      maxWidth: 250,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onHistoryColumnClick(ev, column);
      },
      onRender: (item) => (
        <div>
          <a
            style={{ color: "#0d0091" }}
            data-interception="off"
            target="_blank"
            href={item.FileLink}
            title={item.FileName}
          >{`${
            item.FileName.length > 40
              ? item.FileName.substring(0, 40) + "..."
              : item.FileName
          }`}</a>
        </div>
      ),
    },
    {
      key: "Column4",
      name: "Sent to",
      fieldName: "SentToName",
      minWidth: 75,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onHistoryColumnClick(ev, column);
      },
      onRender: (item) => <div title={item.SentToName}>{item.SentToName}</div>,
    },
    {
      key: "Column5",
      name: "Requests",
      fieldName: "Requests",
      minWidth: 50,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onHistoryColumnClick(ev, column);
      },
    },
    {
      key: "Column6",
      name: "Responses",
      fieldName: "Responses",
      minWidth: 50,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onHistoryColumnClick(ev, column);
      },
    },
    {
      key: "Column7",
      name: "Request comments",
      fieldName: "RequestComments",
      minWidth: 150,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onHistoryColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.RequestComments}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>
            {item.RequestComments.length > 40
              ? item.RequestComments.substring(0, 40) + "..."
              : item.RequestComments}
          </span>
        </TooltipHost>
        // <div title={item.RequestComments}>
        //   {item.RequestComments.length > 40
        //     ? item.RequestComments.substring(0, 40) + "..."
        //     : item.RequestComments}
        // </div>
      ),
    },
    {
      key: "Column8 ",
      name: "Response comments",
      fieldName: "ResponseComments",
      minWidth: 150,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onHistoryColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.ResponseComments}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>
            {item.ResponseComments.length > 40
              ? item.ResponseComments.substring(0, 40) + "..."
              : item.ResponseComments}
          </span>
        </TooltipHost>
        // <div title={item.ResponseComments}>
        // {item.ResponseComments.length > 40
        //   ? item.ResponseComments.substring(0, 40) + "..."
        //   : item.ResponseComments}
        // </div>
      ),
    },
  ];

  const stackTokens: IStackTokens = { childrenGap: 10 };
  const toggleStyles = {
    root: {
      minWidth: 30,
      padding: 0,
      marginRight: 10,
    },
  };
  const statusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: 25,
    width: 100,
  });
  const RequestStyleClass = mergeStyleSets({
    Report: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#895C09",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    Review: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#895C09",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    InitialEdit: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    Assemble: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      statusStyle,
    ],
    AddImages: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      statusStyle,
    ],
    FinalEdit: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    "Sign-off": [
      {
        fontWeight: 600,
        padding: 3,
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      statusStyle,
    ],
    Publish: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      statusStyle,
    ],
  });
  const ResponseStyleClass = mergeStyleSets({
    Pending: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      statusStyle,
    ],
    Cancelled: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      statusStyle,
    ],
    Onhold: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      statusStyle,
    ],
    Feedback: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    Edited: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    Returned: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    SignedOff: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      statusStyle,
    ],
    Completed: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      statusStyle,
    ],
    Assembled: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    Inserted: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    Published: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      statusStyle,
    ],
    Publishready: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      statusStyle,
    ],
    Minorfeedback: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    Majorfeedback: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    Endorsed: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
    Reallocated: [
      {
        fontWeight: 600,
        padding: 3,
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      statusStyle,
    ],
  });

  const DBFilterKeys: IFilter = {
    User: "",
    Week: currentWeekNumber.toString(),
    Year: currentYear.toString(),
    UserState: "Active",
    BA: "All",
    ShowAll: true,
  };
  const DBFilterOptns: IDropdownOptions = {
    User: [{ key: "All", text: "All" }],
    Week: [
      { key: currentWeekNumber.toString(), text: currentWeekNumber.toString() },
    ],
    Year: [{ key: currentYear.toString(), text: currentYear.toString() }],
    UserState: [
      { key: "All", text: "All" },
      { key: "Active", text: "Active" },
      { key: "Inactive", text: "Inactive" },
    ],
    BA: [{ key: "All", text: "All" }],
  };
  // variable-Declaration Ends

  // Style-Declaration Starts
  const projectInfo = mergeStyles({
    display: "flex",
    alignItems: "left",
    justifyContent: "left",
    margin: "0 15px 0 0",
    marginBottom: 10,
  });
  const labelStyles = mergeStyleSets({
    titleLabel: [
      {
        color: "#676767",
        fontSize: 14,
        marginRight: 10,
        fontWeight: 600,
      },
    ],
    labelValue: [
      {
        color: "#0882A5",
        fontSize: 14,
        marginRight: 10,
        fontWeight: 600,
      },
    ],
  });
  const DBfilterLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const DBDropdownStyles: Partial<IDropdownStyles> = {
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
  const DBActiveDropdownStyles: Partial<IDropdownStyles> = {
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
  };
  const DBfilterShortLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 75,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const DBShortDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 75,
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
  const DBActiveShortDropdownStyles: Partial<IDropdownStyles> = {
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
  const DBSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 186,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: 4,
    },
    icon: { fontSize: 12, color: "#000" },
  };
  const DBActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 186,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "2px solid #038387",
      borderRadius: 4,
    },
    field: { fontWeight: 600, color: "#038387" },
    icon: { fontSize: 12, color: "#038387" },
  };
  const DBIconStyleClass = mergeStyleSets({
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
    export: {
      color: "#038387",
      fontSize: "18px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
    },
    backIcon: {
      color: "#000",
      fontSize: "16px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      fontWeight: 600,
      marginTop: 8,
    },
  });
  // Style-Declaration Ends

  // State-Declaration Starts
  const [DBMasterData, setDBMasterData] = useState<IData[]>([]);
  const [DBData, setDBData] = useState<IData[]>([]);
  const [DBDisplayData, setDBDisplayData] = useState<IData[]>([]);
  const [DBFilter, setDBFilter] = useState<IFilter>(DBFilterKeys);
  const [DBFilterDrpDown, setDBFilterDrpDown] =
    useState<IDropdownOptions>(DBFilterOptns);
  const [DBFilterData, setDBFilterData] = useState<IData[]>([]);
  const [DBColumns, setDBColumns] = useState<IColumn[]>(_DBColumns);
  const [DBCurrentPage, setDBCurrentPage] = useState<number>(CurrentPage);
  const [detailHistory, setDetailHistory] = useState<{
    condition: boolean;
    userName: string;
    userEmail: string;
    type: string;
    data: IHistoryData[];
  }>({ condition: false, userName: "", userEmail: "", type: "", data: [] });
  // const [DBHistoryColumns, setDBHistoryColumns] = useState<IColumn[]>(
  //   _DBHistoryColumns
  // );
  const [DBHistoryColumns, setDBHistoryColumns] =
    useState<IColumn[]>(_DBHistoryColumns);
  const [filteredHistoryData, setFilteredHistoryData] = useState<
    IHistoryData[]
  >([]);
  const [DBLoader, setDBLoader] = useState("noLoader");
  const [DBHTML, setDBHTML] = useState("");

  // State-Declaration Ends

  // Function-Declaration Starts

  const queryGenerator = (query): string => {
    let queryStr: string = "";
    if (query.length > 1) {
      var lastTwoStatement = query.length - 2;
      queryStr = "<Where><And>";
      for (let i = 0; i < query.length - 1; i++) {
        if (i == lastTwoStatement) {
          queryStr = queryStr + query[i];
          queryStr = queryStr + query[i + 1];
          break;
        } else {
          queryStr = queryStr + query[i] + "<And>";
        }
      }
      for (let i = 0; i < query.length - 1; i++) {
        queryStr += "</And>";
      }
      queryStr += "</Where>";

      queryStr = queryStr.replace(/\n/g, "");
    } else {
      queryStr = `<Where>`;
      queryStr += query[0];
      queryStr += `</Where>`;

      queryStr = queryStr.replace(/\n/g, "");
    }

    return queryStr;
  };
  const getThresholdData = (
    listName: string,
    filterCondition: string,
    _filterKeys: IFilter,
    weekNumber: number,
    year: number
  ): void => {
    sharepointWeb.lists
      .getByTitle(listName)
      .renderListDataAsStream({
        ViewXml: filterCondition,
      })
      .then((data) => {
        listName == "ProductionBoard"
          ? globalPBData.push(...data.Row)
          : listName == "ActivityProductionBoard"
          ? globalAPBData.push(...data.Row)
          : listName == "Review Log"
          ? globalDRData.push(...data.Row)
          : null;

        if (data.NextHref) {
          getPagedValues(
            listName,
            filterCondition,
            data.NextHref,
            _filterKeys,
            weekNumber,
            year
          );
        } else {
          listName == "ProductionBoard"
            ? getActivityProductionBoardData(_filterKeys, weekNumber, year)
            : listName == "ActivityProductionBoard"
            ? getReviewLogData(_filterKeys, weekNumber, year)
            : listName == "Review Log"
            ? dataManipulationFunction(_filterKeys)
            : null;
        }
      })
      .catch((err: string) => {
        DBErrorFunction(err, `${listName}-getData`);
      });
  };
  const getPagedValues = (
    listName: string,
    filterCondition: string,
    nextHref: string,
    _filterKeys: IFilter,
    weekNumber: number,
    year: number
  ): void => {
    sharepointWeb.lists
      .getByTitle(listName)
      .renderListDataAsStream({
        ViewXml: filterCondition,
        Paging: nextHref.substring(1),
      })
      .then((data) => {
        listName == "ProductionBoard"
          ? globalPBData.push(...data.Row)
          : listName == "ActivityProductionBoard"
          ? globalAPBData.push(...data.Row)
          : listName == "Review Log"
          ? globalDRData.push(...data.Row)
          : null;

        if (data.NextHref) {
          getPagedValues(
            listName,
            filterCondition,
            data.NextHref,
            _filterKeys,
            weekNumber,
            year
          );
        } else {
          listName == "ProductionBoard"
            ? getActivityProductionBoardData(_filterKeys, weekNumber, year)
            : listName == "ActivityProductionBoard"
            ? getReviewLogData(_filterKeys, weekNumber, year)
            : listName == "Review Log"
            ? dataManipulationFunction(_filterKeys)
            : null;
        }
      })
      .catch((err: string) => {
        DBErrorFunction(err, `${listName}-getData`);
      });
  };

  // get Data functions //
  const getMasterUserListData = (
    _filterKeys: IFilter,
    weekNumber: number,
    year: number
  ): void => {
    const sortFilterKeys = (a, b) => {
      if (a.Title < b.Title) {
        return -1;
      }
      if (a.Title > b.Title) {
        return 1;
      }
      return 0;
    };
    setDBLoader("StartLoader");
    globalMasterUserListData = [];

    sharepointWeb.lists
      .getByTitle("Master User List")
      .items.select("*,User/EMail,User/Title")
      .expand("User")
      .filter(`BusinessArea eq '${currentBA}'`)
      .top(5000)
      .get()
      .then((items) => {
        items = items.filter((user) => {
          return user.UserId;
        });
        globalMasterUserListData.push(...items);
        globalMasterUserListData.sort(sortFilterKeys);
        getProductionBoardData(_filterKeys, weekNumber, year);
      })
      .catch((err: string) => {
        DBErrorFunction(err, "getMasterUserListData");
      });
  };
  const getProductionBoardData = (
    _filterKeys: IFilter,
    weekNumber: number,
    year: number
  ): void => {
    globalPBData = [];
    let queryArr = [
      `<Eq>
      <FieldRef Name='Week' />
      <Value Type='Number'>${weekNumber}</Value>
   </Eq>`,
      `<Eq>
      <FieldRef Name='Year' />
      <Value Type='Number'>${year}</Value>
   </Eq>`,
    ];
    let productionBoardQuery = queryGenerator(queryArr);

    let Filtercondition = `
    <View Scope='RecursiveAll'>
      <Query>
         <OrderBy>
           <FieldRef Name='ID' Ascending='FALSE'/>
         </OrderBy>
         ${productionBoardQuery ? productionBoardQuery : null}
      </Query>
      <ViewFields>
        <FieldRef Name='ID' />
        <FieldRef Name='ActualHours' />
        <FieldRef Name='PlannedHours' />
        <FieldRef Name='Developer' />
        <FieldRef Name='Title' />
        <FieldRef Name='EndDate'/>
      </ViewFields>
      <RowLimit Paged='TRUE'>5000</RowLimit>
    </View>`;

    getThresholdData(
      "ProductionBoard",
      Filtercondition,
      _filterKeys,
      weekNumber,
      year
    );
  };
  const getActivityProductionBoardData = (
    _filterKeys: IFilter,
    weekNumber: number,
    year: number
  ): void => {
    globalAPBData = [];
    let queryArr = [
      `<Eq>
      <FieldRef Name='Week' />
      <Value Type='Number'>${weekNumber}</Value>
   </Eq>`,
      `<Eq>
      <FieldRef Name='Year' />
      <Value Type='Number'>${year}</Value>
   </Eq>`,
    ];
    let activityProductionBoardQuery = queryGenerator(queryArr);

    let Filtercondition = `
    <View Scope='RecursiveAll'>
      <Query>
         <OrderBy>
           <FieldRef Name='ID' Ascending='FALSE'/>
         </OrderBy>
         ${activityProductionBoardQuery ? activityProductionBoardQuery : null}
      </Query>
      <ViewFields>
        <FieldRef Name='ID' />
        <FieldRef Name='ActualHours' />
        <FieldRef Name='PlannedHours' />
        <FieldRef Name='Developer' />
        <FieldRef Name='Steps' />
        <FieldRef Name='EndDate'/>
      </ViewFields>
      <RowLimit Paged='TRUE'>5000</RowLimit>
    </View>`;

    getThresholdData(
      "ActivityProductionBoard",
      Filtercondition,
      _filterKeys,
      weekNumber,
      year
    );
  };
  const getReviewLogData = (
    _filterKeys: IFilter,
    weekNumber: number,
    year: number
  ): void => {
    globalDRData = [];
    let dateOfaWeek = moment().isoWeek(weekNumber).year(year);

    let startDateOfaWeek = moment(dateOfaWeek).startOf("week");
    let endDateOfaWeek = moment(dateOfaWeek).endOf("week");

    let queryArr = [
      `<Geq>
      <FieldRef Name='Created' />
      <Value IncludeTimeValue='TRUE' Type='DateTime'>${moment(
        startDateOfaWeek
      ).format("YYYY-MM-DDT00:00:00Z")}</Value>
   </Geq>`,
      `<Leq>
   <FieldRef Name='Created' />
   <Value IncludeTimeValue='TRUE' Type='DateTime'>${moment(
     endDateOfaWeek
   ).format("YYYY-MM-DDT00:00:00Z")}</Value>
</Leq>`,
    ];

    let reviewLogQuery = queryGenerator(queryArr);

    let Filtercondition = `
    <View Scope='RecursiveAll'>
      <Query>
         <OrderBy>
           <FieldRef Name='ID' Ascending='FALSE'/>
         </OrderBy>
         ${reviewLogQuery ? reviewLogQuery : null}
      </Query>
      <ViewFields>
      <FieldRef Name='auditRequestType' />
      <FieldRef Name='auditResponseType' />
      <FieldRef Name='FromUser' />
      <FieldRef Name='auditFrom' />
      <FieldRef Name='ToUser' />
      <FieldRef Name='auditTo' />
      <FieldRef Name='Title' />
      <FieldRef Name='auditLink' />
      <FieldRef Name='auditSent' />
      <FieldRef Name='Modified' />
      <FieldRef Name='Rating' />
      <FieldRef Name='Response_x0020_Comments' />
      <FieldRef Name='auditComments' />
      <FieldRef Name='ProductionBoardID' />
      <FieldRef Name='DRPageName' />
      </ViewFields>
      <RowLimit Paged='TRUE'>5000</RowLimit>
    </View>`;

    getThresholdData(
      "Review Log",
      Filtercondition,
      _filterKeys,
      weekNumber,
      year
    );
  };

  const dataManipulationFunction = (_filterKeys: IFilter): void => {
    let tempMasterData: IData[] = [];

    let validPBData = globalPBData.filter((_PBData) => {
      return _PBData.Developer;
    });
    let validAPBData = globalAPBData.filter((_APBData) => {
      return _APBData.Developer;
    });
    let validDRData = globalDRData.filter((_DRData) => {
      return _DRData.FromUser;
    });

    globalMasterUserListData.forEach((userData: any, index: number) => {
      let tempReviewLog = getFilteredReviewLogData(
        userData.UserId,
        validDRData
      );
      let valAH = calcualteHours(
        "AH",
        userData.UserId,
        validPBData,
        validAPBData
      );
      let valPH = calcualteHours(
        "PH",
        userData.UserId,
        validPBData,
        validAPBData
      );

      tempMasterData.push({
        ID: index,
        UserID: userData.UserId,
        UserName: userData.UserId ? userData.User.Title : "",
        UserEmail: userData.UserId ? userData.User.EMail : "",

        maxProjects: getPBProjectDetails(
          userData.UserId,
          validPBData,
          validAPBData,
          validDRData
        ),

        BA: userData.BusinessArea,
        ActiveStatus: userData.Active ? "Active" : "Inactive",

        // TH: 0,
        TH: userData.SetWeeklyProductionHours
          ? userData.SetWeeklyProductionHours
          : 0,
        AH: valAH,
        PH: valPH,

        ShowAll: Math.abs(valAH - valPH) > 5 || valPH == 0 ? true : false,

        Review: tempReviewLog.Review,
        Edit: tempReviewLog.Edit,
        Assemble: tempReviewLog.Assemble,
        SignOff: tempReviewLog.SignOff,
        Publish: tempReviewLog.Publish,
        Returned: tempReviewLog.Returned,
        Feedback: tempReviewLog.Feedback,
        Actioned: tempReviewLog.Actioned,
        Endosed: tempReviewLog.Endosed,
        SignedOff: tempReviewLog.SignedOff,
        RepeatedIssues: tempReviewLog.RepeatedIssues,
        Major_Moderate: tempReviewLog.Major_Moderate,
        Incomplete: tempReviewLog.Incomplete,
        Quality: tempReviewLog.Quality,
        Style: tempReviewLog.Style,

        ReviewData: [...tempReviewLog.ReviewData],
        EditData: [...tempReviewLog.EditData],
        AssembleData: [...tempReviewLog.AssembleData],
        SignOffData: [...tempReviewLog.SignOffData],
        PublishData: [...tempReviewLog.PublishData],
        ReturnedData: [...tempReviewLog.ReturnedData],
        FeedbackData: [...tempReviewLog.FeedbackData],
        ActionedData: [...tempReviewLog.AssembleData],
        EndosedData: [...tempReviewLog.EndosedData],
        SignedOffData: [...tempReviewLog.SignedOffData],
        RepeatedIssuesData: [...tempReviewLog.RepeatedIssuesData],
        Major_ModerateData: [...tempReviewLog.Major_ModerateData],
        IncompleteData: [...tempReviewLog.IncompleteData],
        QualityData: [...tempReviewLog.QualityData],
        StyleData: [...tempReviewLog.StyleData],
      });
    });
    console.log(tempMasterData, "data");

    filterFunction(tempMasterData, _filterKeys);
    setDBFilter({ ..._filterKeys });

    setDBData([...tempMasterData]);
    sortData = tempMasterData;
    setDBMasterData([...tempMasterData]);
    reloadFilterDropdowns([...tempMasterData]);

    setDBColumns(_DBColumns);
    setDBLoader("noLoader");
  };

  const calcualteHours = (
    type: string,
    userID: number,
    PBData: any,
    APBData: any
  ): number => {
    let tempAH: number = 0;

    let filteredPBData = PBData.filter((_PBData) => {
      return _PBData.Developer[0].id == userID;
    });
    let filteredAPBData = APBData.filter((_APBData) => {
      return _APBData.Developer[0].id == userID;
    });

    if (type == "AH") {
      let PB_AHSum: number =
        filteredPBData.length > 0
          ? filteredPBData.reduce((sum: number, object) => {
              return (
                sum + parseFloat(object.ActualHours ? object.ActualHours : 0)
              );
            }, 0)
          : 0;
      let APB_AHSum: number =
        filteredAPBData.length > 0
          ? filteredAPBData.reduce((sum: number, object) => {
              return (
                sum + parseFloat(object.ActualHours ? object.ActualHours : 0)
              );
            }, 0)
          : 0;

      tempAH = PB_AHSum + APB_AHSum;
    }

    if (type == "PH") {
      let PB_PHSum: number =
        filteredPBData.length > 0
          ? filteredPBData.reduce((sum: number, object) => {
              return (
                sum + parseFloat(object.PlannedHours ? object.PlannedHours : 0)
              );
            }, 0)
          : 0;
      let APB_PHSum: number =
        filteredAPBData.length > 0
          ? filteredAPBData.reduce((sum: number, object) => {
              return (
                sum + parseFloat(object.PlannedHours ? object.PlannedHours : 0)
              );
            }, 0)
          : 0;

      tempAH = PB_PHSum + APB_PHSum;
    }

    return tempAH;
  };
  const getFilteredReviewLogData = (userID: number, DRData: any) => {
    let resultObject = {
      totalCount: 0,

      Review: 0,
      Edit: 0,
      Assemble: 0,
      SignOff: 0,
      Publish: 0,
      Returned: 0,
      Feedback: 0,
      Actioned: 0,
      Endosed: 0,
      SignedOff: 0,
      RepeatedIssues: 0,
      Major_Moderate: 0,
      Incomplete: 0,
      Quality: 0,
      Style: 0,

      ReviewData: [],
      EditData: [],
      AssembleData: [],
      SignOffData: [],
      PublishData: [],
      ReturnedData: [],
      FeedbackData: [],
      ActionedData: [],
      EndosedData: [],
      SignedOffData: [],
      RepeatedIssuesData: [],
      Major_ModerateData: [],
      IncompleteData: [],
      QualityData: [],
      StyleData: [],
    };

    DRData = DRData.filter((_DRData: any) => {
      return _DRData.ProductionBoardID && _DRData.DRPageName;
    });

    let tempFilteredReviewLogdata = DRData.filter((_DRData: any) => {
      return _DRData.FromUser[0].id == userID;
    });

    resultObject.totalCount = tempFilteredReviewLogdata.length;

    // Review
    let FilteredReviewArr = tempFilteredReviewLogdata.filter((obj) => {
      return obj.auditRequestType == "Review";
    });
    resultObject.Review = FilteredReviewArr.length;

    let reviewArr: IHistoryData[] = [];
    if (FilteredReviewArr.length > 0) {
      FilteredReviewArr.forEach((_obj) => {
        reviewArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },

          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });

      resultObject.ReviewData = [...reviewArr];
    }

    // Edit
    let FilteredEditArr = tempFilteredReviewLogdata.filter((obj) => {
      return (
        obj.auditRequestType == "Initial Edit" ||
        obj.auditRequestType == "Final Edit"
      );
    });
    resultObject.Edit = FilteredEditArr.length;

    let editArr: IHistoryData[] = [];
    if (FilteredEditArr.length > 0) {
      FilteredEditArr.forEach((_obj) => {
        editArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });

      resultObject.EditData = [...editArr];
    }

    // Assemble
    let FilteredAssembleArr = tempFilteredReviewLogdata.filter((obj) => {
      return obj.auditRequestType == "Assemble";
    });
    resultObject.Assemble = FilteredAssembleArr.length;

    let assembleArr: IHistoryData[] = [];
    if (FilteredAssembleArr.length > 0) {
      FilteredAssembleArr.forEach((_obj) => {
        assembleArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });
      resultObject.AssembleData = [...assembleArr];
    }

    // SignOff
    let FilteredSignOffArr = tempFilteredReviewLogdata.filter((obj) => {
      return obj.auditRequestType == "Sign-off";
    });
    resultObject.SignOff = FilteredSignOffArr.length;

    let signOffArr: IHistoryData[] = [];
    if (FilteredSignOffArr.length > 0) {
      FilteredSignOffArr.forEach((_obj) => {
        signOffArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });
      resultObject.SignOffData = [...signOffArr];
    }

    // Publish
    let FilteredPublishArr = tempFilteredReviewLogdata.filter((obj) => {
      return obj.auditRequestType == "Publish";
    });
    resultObject.Publish = FilteredPublishArr.length;

    let PublishArr: IHistoryData[] = [];
    if (FilteredPublishArr.length > 0) {
      FilteredPublishArr.forEach((_obj) => {
        PublishArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });
      resultObject.SignOffData = [...signOffArr];
    }

    // Returned
    let FilteredReturnedArr = tempFilteredReviewLogdata.filter((obj) => {
      return obj.auditResponseType == "Returned";
    });
    resultObject.Returned = FilteredReturnedArr.length;

    let ReturnedArr: IHistoryData[] = [];
    if (FilteredReturnedArr.length > 0) {
      FilteredReturnedArr.forEach((_obj) => {
        ReturnedArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });
      resultObject.ReturnedData = [...ReturnedArr];
    }

    // Feedback
    let FilteredFeedbackArr = tempFilteredReviewLogdata.filter((obj) => {
      return (
        obj.auditResponseType == "Minor feedback" ||
        obj.auditResponseType == "Major feedback" ||
        obj.auditResponseType == "Feedback"
      );
    });
    resultObject.Feedback = FilteredFeedbackArr.length;

    let FeedbackArr: IHistoryData[] = [];
    if (FilteredFeedbackArr.length > 0) {
      FilteredFeedbackArr.forEach((_obj) => {
        FeedbackArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });
      resultObject.FeedbackData = [...FeedbackArr];
    }

    // Actioned
    let FilteredActionedArr = tempFilteredReviewLogdata.filter((obj) => {
      return obj.auditResponseType == "Actioned";
    });
    resultObject.Actioned = FilteredActionedArr.length;

    let ActionedArr: IHistoryData[] = [];
    if (FilteredActionedArr.length > 0) {
      FilteredActionedArr.forEach((_obj) => {
        ActionedArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });
      resultObject.ActionedData = [...ActionedArr];
    }

    // Endorsed
    let FilteredEndorsedArr = tempFilteredReviewLogdata.filter((obj) => {
      return obj.auditResponseType == "Endorsed";
    });
    resultObject.Endosed = FilteredEndorsedArr.length;

    let EndorsedArr: IHistoryData[] = [];
    if (FilteredEndorsedArr.length > 0) {
      FilteredEndorsedArr.forEach((_obj) => {
        EndorsedArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });
      resultObject.EndosedData = [...EndorsedArr];
    }

    // SignedOff
    let FilteredSignedOffArr = tempFilteredReviewLogdata.filter((obj) => {
      return obj.auditResponseType == "Signed Off";
    });
    resultObject.SignedOff = FilteredSignedOffArr.length;

    let SignedOffArr: IHistoryData[] = [];
    if (FilteredSignedOffArr.length > 0) {
      FilteredSignedOffArr.forEach((_obj) => {
        SignedOffArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });
      resultObject.SignedOffData = [...SignedOffArr];
    }

    // RepeatedIssues
    let FilteredRepeatedIssuesArr = tempFilteredReviewLogdata.filter((obj) => {
      return obj.FeedbackRepeated == "Yes";
    });
    resultObject.RepeatedIssues = FilteredRepeatedIssuesArr.length;

    let RepeatedIssuesArr: IHistoryData[] = [];
    if (FilteredRepeatedIssuesArr.length > 0) {
      FilteredRepeatedIssuesArr.forEach((_obj) => {
        RepeatedIssuesArr.push({
          FileLink: _obj.auditLink,
          FileName: _obj.Title,
          Sent: _obj["auditSent."],
          SentToName: _obj.ToUser ? _obj.ToUser[0].title : null,

          FromName: _obj.FromUser ? _obj.FromUser[0].title : null,
          ResponseDate: _obj["Modified."],
          Rating: _obj.Rating ? _obj.Rating : null,
          Requests: _obj.auditRequestType,
          Responses: _obj.auditResponseType,
          ResponseComments: _obj.Response_x0020_Comments
            ? _obj.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: _obj.auditComments,

          From: _obj.FromUser
            ? {
                ID: _obj.FromUser[0].id,
                Name: _obj.FromUser[0].title,
                Email: _obj.FromUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          SentTo: _obj.ToUser
            ? {
                ID: _obj.ToUser[0].id,
                Name: _obj.ToUser[0].title,
                Email: _obj.ToUser[0].email,
              }
            : { ID: null, Name: "", Email: "" },
          PBID: _obj.ProductionBoardID,
          PBType:
            _obj.DRPageName == "Annual Plan"
              ? "PB"
              : _obj.DRPageName == "Activity Plan"
              ? "APB"
              : "",
        });
      });
      resultObject.RepeatedIssuesData = [...RepeatedIssuesArr];
    }

    return resultObject;
  };
  const getPBProjectDetails = (
    userID: number,
    PBData: any[],
    APBData: any[],
    DRData: any[]
  ): projectDetails[] => {
    let resArr: projectDetails[] = [];

    let filterPBData = PBData.filter((arr) => {
      return arr.Developer[0].id == userID;
    });

    for (let i = 0; i < filterPBData.length; i++) {
      let arr = filterPBData[i];
      resArr.push({
        ID: arr.ID,
        PBType: "PB",
        projectName: arr.Title,
        AH: arr.ActualHours,
        PH: arr.PlannedHours,
        EndDate: arr.EndDate,
      });
    }

    let filterAPBData = APBData.filter((arr) => {
      return arr.Developer[0].id == userID;
    });
    for (let i = 0; i < filterAPBData.length; i++) {
      let arr = filterAPBData[i];
      resArr.push({
        ID: arr.ID,
        PBType: "APB",
        projectName: arr.Steps,
        AH: arr.ActualHours,
        PH: arr.PlannedHours,
        EndDate: arr.EndDate,
      });
    }

    return resArr;
  };

  const onChangeFilterHandler = (key: string, value: any): void => {
    let tempData: IData[] = DBData;
    let tempFilters: IFilter = DBFilter;
    tempFilters[key] = value;
    setDBFilter({ ...tempFilters });

    if (key == "Week" || key == "Year") {
      setDBLoader("StartLoader");
      getMasterUserListData(
        tempFilters,
        parseInt(tempFilters.Week),
        parseInt(tempFilters.Year)
      );
    } else {
      filterFunction(tempData, tempFilters);
    }
  };
  const filterFunction = (data: IData[], filterKeys: IFilter) => {
    let tempData: IData[] = data;
    let tempFilters: IFilter = filterKeys;
    if (tempFilters.User) {
      tempData = tempData.filter((arr) => {
        return arr.UserName.toLowerCase().includes(
          tempFilters.User.toLowerCase()
        );
      });
    }
    if (tempFilters.UserState != "All") {
      tempData = tempData.filter((arr) => {
        return arr.ActiveStatus == tempFilters.UserState;
      });
      if (tempFilters.UserState == "Active") {
        tempData = tempData.filter((arr) => {
          return arr.UserName.toLowerCase().indexOf("archive") == -1;
        });
      }
    }
    if (tempFilters.BA != "All") {
      tempData = tempData.filter((arr) => {
        return arr.BA == tempFilters.BA;
      });
    }
    if (tempFilters.ShowAll) {
      tempData = tempData.filter((arr) => {
        return arr.ShowAll == true;
      });
    }

    setDBFilterData([...tempData]);
    sortFilterData = tempData;
    paginateFunction(1, tempData);
  };

  const reloadFilterDropdowns = (data: IData[]): void => {
    data.forEach((obj) => {
      if (
        DBFilterOptns.BA.findIndex((BA) => {
          return BA.key == obj.BA;
        }) == -1 &&
        obj.BA
      ) {
        DBFilterOptns.BA.push({
          key: obj.BA,
          text: obj.BA,
        });
      }
    });

    let maxWeek =
      parseInt(DBFilter.Year) == currentYear ? currentWeekNumber : 53;
    for (let i = 1; i <= maxWeek; i++) {
      DBFilterOptns.Week.push({
        key: i.toString(),
        text: i.toString(),
      });
    }
    for (let i = 2020; i <= currentYear; i++) {
      DBFilterOptns.Year.push({
        key: i.toString(),
        text: i.toString(),
      });
    }

    DBFilterOptns.Week.shift();

    DBFilterOptns.Year.shift();

    setDBFilterDrpDown(DBFilterOptns);
  };

  // column-sorting function //
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempORColumns = _DBColumns;
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
      sortData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newORFilterData = _copyAndSort(
      sortFilterData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setDBData([...newORData]);
    setDBFilterData([...newORFilterData]);
    paginateFunction(1, [...newORFilterData]);
  };
  const _onHistoryColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempColumns = _DBHistoryColumns;
    const newColumns: IColumn[] = tempColumns.slice();
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
      sortHistoryData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setFilteredHistoryData([...newORData]);
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

  const generateExcel = (): void => {
    let arrExport = DBFilterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");

    worksheet.columns = [
      { header: "User", key: "User", width: 25 },
      { header: "AH", key: "AH", width: 25 },
      { header: "PH", key: "PH", width: 25 },
      { header: "Review", key: "Review", width: 25 },
      { header: "Edit", key: "Edit", width: 25 },
      { header: "Assemble", key: "Assemble", width: 25 },
      { header: "SignOff", key: "SignOff", width: 25 },
      { header: "Returned", key: "Returned", width: 25 },
      { header: "Feedback", key: "Feedback", width: 25 },
      { header: "Actioned", key: "Actioned", width: 25 },
      { header: "Endosed", key: "Endosed", width: 25 },
      { header: "SignedOff", key: "SignedOff", width: 25 },
      { header: "RepeatedIssues", key: "RepeatedIssues", width: 25 },
    ];

    /*arrExport.forEach((item: IData) => {
      worksheet.addRow({
        User: item.UserName ? item.UserName : "",
        AH: item.AH ? item.AH : 0,
        PH: item.PH ? item.PH : 0,
        Review: item.Review ? item.Review : 0,
        Edit: item.Edit ? item.Edit : 0,
        Assemble: item.Assemble ? item.Assemble : 0,
        SignOff: item.SignOff ? item.SignOff : 0,
        Returned: item.Returned ? item.Returned : 0,
        Feedback: item.Feedback ? item.Feedback : 0,
        Actioned: item.Actioned ? item.Actioned : 0,
        Endosed: item.Endosed ? item.Endosed : 0,
        SignedOff: item.SignedOff ? item.SignedOff : 0,
        RepeatedIssues: item.RepeatedIssues ? item.RepeatedIssues : 0,
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
    ].map((key) => {
      worksheet.getCell(key).color = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF" },
      };
    });*/
    workbook.xlsx
      .writeBuffer()
      .then((buffer) =>
        FileSaver.saveAs(
          new Blob([buffer]),
          `Weeklyreport-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };

  const exportToExcel = () => {
    var location = "data:application/vnd.ms-excel;base64,";
    let designHtml = `<style>
    th  
{  

    background-color :grey;  
    color: white;  
      
}  
    </style>`;

    var excelTemplate =
      "<html> " +
      "<head> " +
      '<meta http-equiv="content-type" content="text/plain; charset=UTF-8"/> ' +
      "</head> " +
      designHtml +
      "<body> " +
      DBHTML +
      "</body> " +
      "</html>";
    console.log(excelTemplate);
    var uri = "data:application/vnd.ms-excel;base64,",
      template =
        '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>',
      base64 = function (s) {
        return window.btoa(unescape(encodeURIComponent(s)));
      },
      format = function (s, c) {
        return s.replace(/{(\w+)}/g, function (m, p) {
          return c[p];
        });
      };
    // if (!table.nodeType) table = document.getElementById(table)
    var ctx = { worksheet: "Sheet1" || "Worksheet", table: excelTemplate };
    var a = document.createElement("a");
    a.href = uri + base64(format(template, ctx));
    a.download = `Productionreport-${new Date().toLocaleString()}.xls`;
    //triggering the function
    a.click();
  };

  const paginateFunction = (pagenumber, data): void => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      CurrentPage = pagenumber;
      setDBDisplayData(paginatedItems);
      setDBCurrentPage(pagenumber);
    } else {
      setDBDisplayData([]);
      setDBCurrentPage(1);
    }
  };

  const getHTMLFunction = (html: string) => {
    setDBHTML(html);
  };

  const pageSwitchFunction = (
    _condition: boolean,
    _userName: string,
    _userEmail: string,
    _type: string,
    _data: IHistoryData[]
  ) => {
    setDetailHistory({
      condition: _condition,
      userName: _userName,
      userEmail: _userEmail,
      type: _type,
      data: _data,
    });
  };

  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  const DBErrorFunction = (error: any, functionName: string): void => {
    console.log(error, functionName);

    let response = {
      ComponentName: "Weekly report - dashboard",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setDBLoader("noLoader");
        ErrorPopup();
      }
    );
  };

  // Function-Declaration Ends

  useEffect(() => {
    getMasterUserListData(DBFilterKeys, currentWeekNumber, currentYear);
  }, [currentBA]);

  // return (
  //   <div>
  //     <ProductReport Data={DBData} />
  //   </div>
  // );

  return DBMasterData ? (
    <>
      {detailHistory.condition ? (
        <div>
          <div style={{ padding: 10, marginTop: 10 }}>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                marginBottom: 10,
              }}
            >
              <Icon
                iconName="ChromeBack"
                className={DBIconStyleClass.backIcon}
                onClick={(): void => {
                  setDBHistoryColumns(_DBHistoryColumns);
                  setDetailHistory({
                    condition: false,
                    userEmail: "",
                    userName: "",
                    type: "",
                    data: [],
                  });
                  setFilteredHistoryData([]);
                  sortHistoryData = [];
                }}
              />
              <div
                style={{
                  display: "flex",
                }}
              >
                <div style={{ cursor: "pointer" }}>
                  <Persona
                    title={detailHistory.userName}
                    size={PersonaSize.size32}
                    presence={PersonaPresence.none}
                    imageUrl={
                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                      `${detailHistory.userEmail}`
                    }
                  />
                </div>
                <Label
                  styles={{
                    root: {
                      color: "#000",
                      fontSize: "18px",
                      marginLeft: "10px",
                    },
                  }}
                  className={styles.DBHistoryHeading}
                >{`${detailHistory.userName} - ${
                  detailHistory.type.split(",")[0]
                }`}</Label>
              </div>
            </div>
            <div style={{ display: "flex" }}>
              <div className={projectInfo}>
                <Label className={labelStyles.titleLabel}>
                  Name of deliverable or steps:
                </Label>
                <Label className={labelStyles.labelValue}>
                  {detailHistory.type.split(",")[1]}
                </Label>
              </div>
              <div className={projectInfo}>
                <Label className={labelStyles.titleLabel}>Planned date :</Label>
                <Label className={labelStyles.labelValue}>
                  {detailHistory.type.split(",")[2]}
                </Label>
              </div>
            </div>
            <div>
              <DetailsList
                items={detailHistory.data}
                columns={DBHistoryColumns}
                styles={
                  detailHistory.data.length > 0
                    ? {
                        root: {
                          ".ms-DetailsRow-cell": {
                            height: 40,
                          },
                          ".ms-DetailsHeader-cellTitle": {
                            background: "#03828711 !important",
                            color: "#038387 !important",
                          },
                        },
                      }
                    : null
                }
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
              />
              {detailHistory.data.length > 0 ? null : (
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    marginTop: "15px",
                  }}
                >
                  <Label style={{ color: "#2392B2", fontWeight: 600 }}>
                    No data found !!!
                  </Label>
                </div>
              )}
            </div>
          </div>
        </div>
      ) : (
        <div>
          {DBLoader == "StartLoader" ? (
            <CustomLoader />
          ) : (
            <div>
              {/* Header-Section Starts */}
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  marginTop: "10px",
                  flexWrap: "wrap",
                }}
              >
                {/* Filter-Section Starts */}
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    marginBottom: 10,
                    flexWrap: "wrap",
                  }}
                >
                  <div>
                    <Label styles={DBfilterLabelStyles}>User</Label>

                    <SearchBox
                      placeholder="Search user"
                      styles={
                        DBFilter.User
                          ? DBActiveSearchBoxStyles
                          : DBSearchBoxStyles
                      }
                      value={DBFilter.User}
                      onChange={(e, value): void => {
                        onChangeFilterHandler("User", value);
                      }}
                    />
                  </div>
                  <div>
                    <Label styles={DBfilterLabelStyles}>User state</Label>
                    <Dropdown
                      placeholder="Select an option"
                      options={DBFilterDrpDown.UserState}
                      selectedKey={DBFilter.UserState}
                      styles={
                        DBFilter.UserState == "All"
                          ? DBDropdownStyles
                          : DBActiveDropdownStyles
                      }
                      onChange={(e, option: any) => {
                        onChangeFilterHandler("UserState", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label styles={DBfilterShortLabelStyles}>Week</Label>
                    <Dropdown
                      placeholder="Select an option"
                      options={DBFilterDrpDown.Week}
                      selectedKey={DBFilter.Week}
                      styles={
                        DBFilter.Week
                          ? DBActiveShortDropdownStyles
                          : DBShortDropdownStyles
                      }
                      onChange={(e, option: any) => {
                        onChangeFilterHandler("Week", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label styles={DBfilterShortLabelStyles}>Year</Label>
                    <Dropdown
                      placeholder="Select an option"
                      options={DBFilterDrpDown.Year}
                      selectedKey={DBFilter.Year}
                      styles={
                        DBFilter.Year
                          ? DBActiveShortDropdownStyles
                          : DBShortDropdownStyles
                      }
                      onChange={(e, option: any) => {
                        onChangeFilterHandler("Year", option["key"]);
                      }}
                    />
                  </div>
                  <div style={{ marginLeft: "10px", marginRight: "10px" }}>
                    <Stack tokens={stackTokens}>
                      <Toggle
                        label="Show all"
                        styles={toggleStyles}
                        checked={!DBFilter.ShowAll}
                        onChange={(e) => {
                          onChangeFilterHandler("ShowAll", !DBFilter.ShowAll);
                        }}
                      />
                    </Stack>
                  </div>
                  <div>
                    <Icon
                      iconName="Refresh"
                      title="Click to reset"
                      className={DBIconStyleClass.refresh}
                      onClick={() => {
                        getMasterUserListData(
                          DBFilterKeys,
                          currentWeekNumber,
                          currentYear
                        );
                      }}
                    />
                  </div>
                </div>
                {/* Filter-Section Ends */}
                {/* Header-Btn-Section Starts */}
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "left",
                    paddingTop: 16,
                  }}
                >
                  <Label
                    style={{
                      color: "#323130",
                      fontSize: "13px",
                      marginLeft: "10px",
                      fontWeight: "500",
                      marginRight: 10,
                    }}
                  >
                    Number of records:{" "}
                    <b style={{ color: "#038387" }}>{DBFilterData.length}</b>
                  </Label>
                  <Label
                    onClick={() => {
                      exportToExcel();
                      // generateExcel();
                    }}
                    style={{
                      backgroundColor: "#EBEBEB",
                      padding: "7px 15px",
                      cursor: "pointer",
                      fontSize: 12,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      borderRadius: 3,
                      color: "#1D6F42",
                    }}
                  >
                    <Icon
                      style={{
                        color: "#1D6F42",
                      }}
                      iconName="ExcelDocument"
                      className={DBIconStyleClass.export}
                    />
                    Export as XLS
                  </Label>
                </div>
                {/* Header-Btn-Section Ends */}
              </div>
              {/* Header-Section Ends */}

              {/* badge info section starts */}
              {/* <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                }}
              >
                <div className={styles.badgeInfoSection}>
                  <div className={styles.info}>
                    <span style={{ backgroundColor: "#038387" }}></span>Doc
                    review requests
                  </div>
                  <div className={styles.info}>
                    <span style={{ backgroundColor: "#FAA332" }}></span>Doc
                    review response
                  </div>
                  <div className={styles.info}>
                    <span style={{ backgroundColor: "#FA6232" }}></span>Doc
                    quality/feedback
                  </div>
                </div>
              </div> */}
              {/* badge info section Ends */}

              {/* Body-Section Starts */}
              <div>
                {/* DetailList-Section Starts */}
                <PHDashboardTable
                  masterData={...DBMasterData}
                  filteredData={...DBFilterData}
                  displayData={...DBDisplayData}
                  pageSwitch={pageSwitchFunction}
                  getHtml={getHTMLFunction}
                />
                {/* <DetailsList
                  items={DBDisplayData}
                  columns={DBColumns}
                  styles={{
                    root: {
                      ".ms-DetailsRow-cell": {
                        // display: "flex",
                        // alignItems: "center",
                        height: 40,
                      },
                      ".ms-DetailsHeader-cellTitle": {
                        background: "transparent !important",
                      },
                    },
                  }}
                  setKey="set"
                  layoutMode={DetailsListLayoutMode.justified}
                  selectionMode={SelectionMode.none}
                /> */}
                {/* DetailList-Section Ends */}
              </div>
              {DBFilterData.length > 0 ? (
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    margin: "10px 0",
                  }}
                >
                  <Pagination
                    currentPage={DBCurrentPage}
                    totalPages={
                      DBFilterData.length > 0
                        ? Math.ceil(DBFilterData.length / totalPageItems)
                        : 1
                    }
                    onChange={(page) => {
                      paginateFunction(page, DBFilterData);
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
                  <Label style={{ color: "#2392B2", fontWeight: 600 }}>
                    No data Found !!!
                  </Label>
                </div>
              )}
              {/* Body-Section Ends */}
            </div>
          )}
        </div>
      )}
    </>
  ) : null;
};

export default WRDashboard;
