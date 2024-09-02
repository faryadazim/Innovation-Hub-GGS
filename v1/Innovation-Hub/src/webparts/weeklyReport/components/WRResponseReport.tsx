import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
import * as moment from "moment";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
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
  IColumn,
  ILabelStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  Stack,
  IStackTokens,
  Toggle,
  Rating,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./WeeklyReport.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
import CustomLoader from "./CustomLoader";
import Pagination from "office-ui-fabric-react-pagination";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

interface IProps {
  context: WebPartContext;
  spcontext: any;
  graphContent: any;
  URL: string;
  peopleList: any;
  historyDataHandler: any;
  BA: string;
}
interface IFilter {
  from: string;
  to: string;
  requests: string;
  responses: string;
  week: string;
  year: string;
  showAll: boolean;
}
interface IMasterUserListData {
  userID: number;
  userName: string;
  userEmail: string;
  userBA: string;
}
interface IData {
  ID: number;

  FromUserID: number;
  FromUserName: string;
  FromUserEmail: string;

  SentDate: string;
  ResponseDate: string;

  Title: string;
  fileUrl: string;

  ToUserID: number;
  ToUserName: string;
  ToUserEmail: string;

  Rating: number;
  Requests: string;
  Responses: string;
  ResponseComments: string;
  RequestComments: string;

  showAllFlag: boolean;
}
interface IDropdown {
  key: string;
  text: string;
}
interface IDropdownOptions {
  requestsOptns: IDropdown[];
  responsesOptns: IDropdown[];
  weekOptns: IDropdown[];
  yearOptns: IDropdown[];
}

let sortData: IData[] = [];
let sortFilterData: IData[] = [];

let globalMasterUserListData: IMasterUserListData[] = [];
let globalDRData = [];

let CurrentPage: number = 1;
let totalPageItems: number = 10;

const WRReviewReport = (props: IProps) => {
  // variable-Declaration Starts
  const sharepointWeb: any = Web(props.URL);
  const allPeoples: any[] = props.peopleList;
  const currentBA: string = props.BA;

  const currentYear: number = moment().year();
  const currentWeekNumber: number = moment().isoWeek();

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  const _resReportColumns: IColumn[] = [
    {
      key: "Column1",
      name: "File responding",
      fieldName: "ToUserName",
      minWidth: 150,
      maxWidth: 250,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{
            display: "flex",
          }}
        >
          <div style={{ cursor: "pointer" }}>
            <Persona
              title={item.ToUserName}
              size={PersonaSize.size24}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.ToUserEmail}`
              }
            />
          </div>
          <Label style={{ marginTop: -3 }} title={item.ToUserName}>
            {item.ToUserName}
          </Label>
        </div>
      ),
    },
    {
      key: "Column2",
      name: "Title",
      fieldName: "Title",
      minWidth: 150,
      maxWidth: 250,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div>
          <a
            style={{ color: "#0d0091" }}
            data-interception="off"
            target="_blank"
            href={item.fileUrl}
            title={item.Title}
          >{`${
            item.Title.length > 40
              ? item.Title.substring(0, 40) + "..."
              : item.Title
          }`}</a>
        </div>
      ),
    },
    {
      key: "Column3",
      name: "Sent date",
      fieldName: "SentDate",
      minWidth: 80,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div>{moment(item.SentDate).format("DD/MM/yyyy")}</div>
      ),
    },
    {
      key: "Column4",
      name: "Response date",
      fieldName: "ResponseDate",
      minWidth: 110,
      maxWidth: 120,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div>{moment(item.ResponseDate).format("DD/MM/yyyy")}</div>
      ),
    },
    {
      key: "Column5",
      name: "From",
      fieldName: "FromUserName",
      minWidth: 150,
      maxWidth: 250,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{
            display: "flex",
          }}
        >
          <div style={{ cursor: "pointer" }}>
            <Persona
              title={item.FromUserName}
              size={PersonaSize.size24}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.FromUserEmail}`
              }
            />
          </div>
          <Label style={{ marginTop: -3 }} title={item.FromUserName}>
            {item.FromUserName}
          </Label>
        </div>
      ),
    },

    {
      key: "Column6",
      name: "Rating",
      fieldName: "Rating",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div>
          <Rating
            max={4}
            allowZeroStars
            rating={item.Rating}
            readOnly={true}
            styles={{
              ratingStarFront: {
                color:
                  item.Rating == 1
                    ? "#D10000"
                    : item.Rating == 2
                    ? "#D18700"
                    : item.Rating == 3
                    ? "#a3a300"
                    : item.Rating == 4
                    ? "#00a300"
                    : "#038387",
              },
            }}
          />
        </div>
      ),
    },
    {
      key: "Column7",
      name: "Requests",
      fieldName: "Requests",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <>
          <div
            className={RequestStyleClass[`${item.Requests.replace(" ", "")}`]}
          >
            {item.Requests}
          </div>
        </>
      ),
    },
    {
      key: "Column8",
      name: "Responses",
      fieldName: "Responses",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          className={ResponseStyleClass[`${item.Responses.replace(" ", "")}`]}
        >
          {item.Responses}
        </div>
      ),
    },
    {
      key: "Column9",
      name: "Response comments",
      fieldName: "ResponseComments",
      minWidth: 250,
      maxWidth: 400,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div
          style={{ cursor: item.ResponseComments ? "pointer" : "default" }}
          title={item.ResponseComments}
        >
          {item.ResponseComments.length > 40
            ? item.ResponseComments.substring(0, 40) + "..."
            : item.ResponseComments}
        </div>
      ),
    },
    {
      key: "Column10",
      name: "Action",
      fieldName: "Action",
      minWidth: 60,
      maxWidth: 100,

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
              iconName="DocumentReply"
              className={resReportIconStyleClass.historyIcon}
              onClick={(): void => {
                props.historyDataHandler(true, item.ID);
                // getOrgReportHistoryData(item);
              }}
            />
          </div>
        </div>
      ),
    },
  ];
  const resReportFilterKeys: IFilter = {
    from: "",
    to: "",
    requests: "All",
    responses: "All",
    week: currentWeekNumber.toString(),
    year: currentYear.toString(),
    showAll: false,
  };
  const resReportFilterOptns: IDropdownOptions = {
    requestsOptns: [{ key: "All", text: "All" }],
    responsesOptns: [{ key: "All", text: "All" }],
    weekOptns: [
      { key: currentWeekNumber.toString(), text: currentWeekNumber.toString() },
    ],
    yearOptns: [{ key: currentYear.toString(), text: currentYear.toString() }],
  };
  // variable-Declaration Ends

  // Style-Declaration Starts
  const stackTokens: IStackTokens = { childrenGap: 10 };
  const resReportfilterLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const resReportDropdownStyles: Partial<IDropdownStyles> = {
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
  const resReportActiveDropdownStyles: Partial<IDropdownStyles> = {
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
  const resReportfilterShortLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 75,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const resReportShortDropdownStyles: Partial<IDropdownStyles> = {
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
  const resReportActiveShortDropdownStyles: Partial<IDropdownStyles> = {
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
  const resReportSearchBoxStyles: Partial<ISearchBoxStyles> = {
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
  const resReportActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
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
  const resReportIconStyleClass = mergeStyleSets({
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
    historyIcon: {
      color: "#038387",
      fontSize: 20,
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      fontWeight: 600,
    },
  });
  // Style-Declaration Ends

  // State-Declaration Starts
  const [resReportMasterData, setResReportMasterData] = useState<IData[]>([]);
  const [resReportData, setResReportData] = useState<IData[]>([]);
  const [resReportDisplayData, setResReportDisplayData] = useState<IData[]>([]);
  const [resReportFilter, setResReportFilter] =
    useState<IFilter>(resReportFilterKeys);
  const [resReportFilterDrpDown, setResReportFilterDrpDown] =
    useState<IDropdownOptions>(resReportFilterOptns);
  const [resReportFilterData, setResReportFilterData] = useState<IData[]>([]);
  const [resReportColumns, setResReportColumns] =
    useState<IColumn[]>(_resReportColumns);
  const [resReportCurrentPage, setResReportCurrentPage] =
    useState<number>(CurrentPage);
  const [resReportLoader, setResReportLoader] = useState("noLoader");
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
        globalDRData.push(...data.Row);

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
          dataManipulationFunction(_filterKeys);
        }
      })
      .catch((err: string) => {
        resReportErrorFunction(err, `${listName}-getData`);
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
        globalDRData.push(...data.Row);

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
          dataManipulationFunction(_filterKeys);
        }
      })
      .catch((err: string) => {
        resReportErrorFunction(err, `${listName}-getData`);
      });
  };

  const onChangeFilterHandler = (
    key: string,
    value: string | boolean
  ): void => {
    let tempData: IData[] = resReportData;
    let tempFilters: IFilter = resReportFilter;
    tempFilters[key] = value;
    setResReportFilter({ ...tempFilters });

    if (key == "week" || key == "year") {
      getMasterUserListData(
        tempFilters,
        parseInt(tempFilters.week),
        parseInt(tempFilters.year)
      );
    } else {
      filterFunction(tempData, tempFilters);
    }
  };
  const filterFunction = (data: IData[], filterKeys: IFilter) => {
    let fitlerData: IData[] = data.filter((_data) => {
      return _data.Rating < 3;
    });
    let tempData: IData[] = filterKeys.showAll ? data : fitlerData;
    let tempFilters: IFilter = filterKeys;

    if (tempFilters.from) {
      tempData = tempData.filter((arr) => {
        return arr.FromUserName.toLowerCase().includes(
          tempFilters.from.toLowerCase()
        );
      });
    }

    if (tempFilters.to) {
      tempData = tempData.filter((arr) => {
        return arr.ToUserName.toLowerCase().includes(
          tempFilters.to.toLowerCase()
        );
      });
    }

    if (tempFilters.requests != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Requests == tempFilters.requests;
      });
    }

    if (tempFilters.responses != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Responses == tempFilters.responses;
      });
    }

    setResReportFilterData([...tempData]);
    sortFilterData = tempData;
    paginateFunction(1, tempData);
  };
  const reloadFilterDropdowns = (data: IData[]): void => {
    data.forEach((obj) => {
      if (
        resReportFilterOptns.requestsOptns.findIndex((BA) => {
          return BA.key == obj.Requests;
        }) == -1 &&
        obj.Requests
      ) {
        resReportFilterOptns.requestsOptns.push({
          key: obj.Requests,
          text: obj.Requests,
        });
      }

      if (
        resReportFilterOptns.responsesOptns.findIndex((BA) => {
          return BA.key == obj.Responses;
        }) == -1 &&
        obj.Responses
      ) {
        resReportFilterOptns.responsesOptns.push({
          key: obj.Responses,
          text: obj.Responses,
        });
      }
    });

    let maxWeek =
      parseInt(resReportFilter.year) == currentYear ? currentWeekNumber : 53;
    for (let i = 1; i <= maxWeek; i++) {
      resReportFilterOptns.weekOptns.push({
        key: i.toString(),
        text: i.toString(),
      });
    }
    for (let j = 2020; j <= currentYear; j++) {
      resReportFilterOptns.yearOptns.push({
        key: j.toString(),
        text: j.toString(),
      });
    }

    resReportFilterOptns.weekOptns.shift();
    resReportFilterOptns.yearOptns.shift();

    setResReportFilterDrpDown(resReportFilterOptns);
  };

  //get-Data function //
  const getMasterUserListData = (
    _filterKeys: IFilter,
    weekNumber: number,
    year: number
  ): void => {
    setResReportLoader("StartLoader");

    sharepointWeb.lists
      .getByTitle("Master User List")
      .items.select("*,User/ID,User/EMail,User/Title")
      .expand("User")
      .filter(`BusinessArea eq '${currentBA}'`)
      .top(5000)
      .get()
      .then((items) => {
        globalMasterUserListData = [];

        items = items.filter((user) => {
          return user.UserId;
        });

        items.forEach((user) => {
          globalMasterUserListData.push({
            userID: user.User.ID,
            userName: user.User.Title,
            userEmail: user.User.EMail,
            userBA: user.BusinessArea,
          });
        });

        getReviewLogData(_filterKeys, weekNumber, year);
      })
      .catch((err) => {
        resReportErrorFunction(err, "MasterUserListData-getData");
      });
  };
  const getReviewLogData = (
    _filterKeys: IFilter,
    weekNumber: number,
    year: number
  ): void => {
    // setResReportLoader("StartLoader");
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
    // console.log(reviewLogQuery);

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
    console.log(globalMasterUserListData);
    let tempMasterData: IData[] = [];
    globalDRData.forEach((data: any) => {
      if (
        data.ToUser &&
        globalMasterUserListData.some(
          (_user) => _user.userID == data.ToUser[0].id
        )
      ) {
        tempMasterData.push({
          ID: data.ID,

          FromUserID: data.FromUser ? data.FromUser[0].id : null,
          FromUserName: data.FromUser ? data.FromUser[0].title : "",
          FromUserEmail: data.FromUser ? data.FromUser[0].email : "",

          SentDate: data["auditSent."],
          ResponseDate: data["Modified."],

          Title: data.Title,
          fileUrl: data.auditLink ? data.auditLink : "",

          ToUserID: data.FromUser ? data.ToUser[0].id : null,
          ToUserName: data.FromUser ? data.ToUser[0].title : "",
          ToUserEmail: data.FromUser ? data.ToUser[0].email : "",

          Rating: data.Rating ? data.Rating : null,
          Requests: data.auditRequestType,
          Responses: data.auditResponseType,

          ResponseComments: data.Response_x0020_Comments
            ? data.Response_x0020_Comments.replace(/<[^>]+>/g, "")
            : "",
          RequestComments: data.auditComments,

          showAllFlag: data.Rating < 3 ? true : false,
        });
      }
    });

    console.log(tempMasterData);

    filterFunction(tempMasterData, _filterKeys);
    setResReportFilter({ ..._filterKeys });

    setResReportData([...tempMasterData]);
    sortData = tempMasterData;
    setResReportMasterData([...tempMasterData]);
    reloadFilterDropdowns([...tempMasterData]);

    setResReportColumns(_resReportColumns);
    setResReportLoader("noLoader");
  };

  // column-sorting function //
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempORColumns = _resReportColumns;
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
    setResReportData([...newORData]);
    setResReportFilterData([...newORFilterData]);
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

  const generateExcel = (): void => {
    let arrExport = resReportFilterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "File Submitted", key: "FileSubmitted", width: 25 },
      { header: "Sent Day", key: "SentDay", width: 25 },
      { header: "Title", key: "Title", width: 70 },
      { header: "Send To", key: "SendTo", width: 20 },
      { header: "Rating", key: "Rating", width: 25 },
      { header: "Request", key: "Request", width: 25 },
      { header: "Response", key: "Response", width: 25 },
      { header: "Response Comments", key: "ResponseComments", width: 100 },
    ];
    arrExport.forEach((item: IData) => {
      worksheet.addRow({
        FileSubmitted: item.FromUserName ? item.FromUserName : "",
        SentDay: item.SentDate ? moment(item.SentDate).format("dddd") : "",
        Title: item.Title ? item.Title : "",
        SendTo: item.ToUserName ? item.ToUserName : "",
        Rating: item.Rating ? item.Rating : "",
        Request: item.Requests ? item.Requests : "",
        Response: item.Responses ? item.Responses : "",
        ResponseComments: item.ResponseComments ? item.ResponseComments : "",
      });
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"].map((key) => {
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
          `Weeklyreport-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };
  const paginateFunction = (pagenumber, data): void => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      CurrentPage = pagenumber;
      setResReportDisplayData(paginatedItems);
      setResReportCurrentPage(pagenumber);
    } else {
      setResReportDisplayData([]);
      setResReportCurrentPage(1);
    }
  };

  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  const resReportErrorFunction = (error: any, functionName: string): void => {
    console.log(error, functionName);
    let response = {
      ComponentName: "Weekly report - response report",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setResReportLoader("noLoader");
        ErrorPopup();
      }
    );
  };
  // Function-Declaration Ends

  useEffect(() => {
    getMasterUserListData(resReportFilterKeys, currentWeekNumber, currentYear);
  }, [currentBA]);
  return (
    <div>
      {resReportLoader == "StartLoader" ? (
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
                <Label styles={resReportfilterLabelStyles}>From</Label>
                <SearchBox
                  placeholder="Search from user"
                  styles={
                    resReportFilter.from
                      ? resReportActiveSearchBoxStyles
                      : resReportSearchBoxStyles
                  }
                  value={resReportFilter.from}
                  onChange={(e, value): void => {
                    onChangeFilterHandler("from", value);
                  }}
                />
              </div>
              <div>
                <Label styles={resReportfilterLabelStyles}>To</Label>
                <SearchBox
                  placeholder="Search to user"
                  styles={
                    resReportFilter.to
                      ? resReportActiveSearchBoxStyles
                      : resReportSearchBoxStyles
                  }
                  value={resReportFilter.to}
                  onChange={(e, value): void => {
                    onChangeFilterHandler("to", value);
                  }}
                />
              </div>
              <div>
                <Label styles={resReportfilterLabelStyles}>Requests</Label>
                <Dropdown
                  placeholder="Select an option"
                  options={resReportFilterDrpDown.requestsOptns}
                  selectedKey={resReportFilter.requests}
                  styles={
                    resReportFilter.requests != "All"
                      ? resReportActiveDropdownStyles
                      : resReportDropdownStyles
                  }
                  onChange={(e, option: any) => {
                    onChangeFilterHandler("requests", option["key"]);
                  }}
                />
              </div>
              <div>
                <Label styles={resReportfilterLabelStyles}>Responses</Label>
                <Dropdown
                  placeholder="Select an option"
                  options={resReportFilterDrpDown.responsesOptns}
                  selectedKey={resReportFilter.responses}
                  styles={
                    resReportFilter.responses != "All"
                      ? resReportActiveDropdownStyles
                      : resReportDropdownStyles
                  }
                  onChange={(e, option: any) => {
                    onChangeFilterHandler("responses", option["key"]);
                  }}
                />
              </div>
              <div>
                <Label styles={resReportfilterShortLabelStyles}>Week</Label>
                <Dropdown
                  placeholder="Select an option"
                  options={resReportFilterDrpDown.weekOptns}
                  selectedKey={resReportFilter.week}
                  styles={
                    resReportFilter.week
                      ? resReportActiveShortDropdownStyles
                      : resReportShortDropdownStyles
                  }
                  onChange={(e, option: any) => {
                    onChangeFilterHandler("week", option["key"]);
                  }}
                />
              </div>
              <div>
                <Label styles={resReportfilterShortLabelStyles}>Year</Label>
                <Dropdown
                  placeholder="Select an option"
                  options={resReportFilterDrpDown.yearOptns}
                  selectedKey={resReportFilter.year}
                  styles={
                    resReportFilter.year
                      ? resReportActiveShortDropdownStyles
                      : resReportShortDropdownStyles
                  }
                  onChange={(e, option: any) => {
                    onChangeFilterHandler("year", option["key"]);
                  }}
                />
              </div>
              <div style={{ marginLeft: "10px", marginRight: "10px" }}>
                <Stack tokens={stackTokens}>
                  <Toggle
                    label="Show all"
                    styles={toggleStyles}
                    checked={resReportFilter.showAll}
                    onChange={(e) => {
                      onChangeFilterHandler(
                        "showAll",
                        !resReportFilter.showAll
                      );
                    }}
                  />
                </Stack>
              </div>
              <div>
                <Icon
                  iconName="Refresh"
                  title="Click to reset"
                  className={resReportIconStyleClass.refresh}
                  onClick={() => {
                    getMasterUserListData(
                      resReportFilterKeys,
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
                paddingTop: 26,
                paddingBottom: "10px",
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
                <b style={{ color: "#038387" }}>{resReportFilterData.length}</b>
              </Label>
              <Label
                onClick={() => {
                  generateExcel();
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
                  className={resReportIconStyleClass.export}
                />
                Export as XLS
              </Label>
            </div>
            {/* Header-Btn-Section Ends */}
          </div>
          {/* Header-Section Ends */}
          {/* Body-Section Starts */}
          <div>
            {/* DetailList-Section Starts */}
            <DetailsList
              items={resReportDisplayData}
              columns={resReportColumns}
              styles={{
                root: {
                  ".ms-DetailsRow-cell": {
                    height: 40,
                  },
                  ".ms-DetailsHeader-cellTitle": {
                    background: "#03828711 !important",
                    color: "#038387 !important",
                  },
                },
              }}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              onRenderRow={(data, defaultRender) => (
                <div>
                  {defaultRender({
                    ...data,
                    styles: {
                      root: {
                        background:
                          resReportFilter.showAll &&
                          data.item.showAllFlag == true
                            ? "#FFF2F2"
                            : "#fff",
                        selectors: {
                          "&:hover": {
                            background:
                              resReportFilter.showAll &&
                              data.item.showAllFlag == true
                                ? "#f5e3e3"
                                : "#f3f2f1",
                          },
                        },
                      },
                    },
                  })}
                </div>
              )}
            />
            {/* DetailList-Section Ends */}
          </div>
          {resReportFilterData.length > 0 ? (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                margin: "10px 0",
              }}
            >
              <Pagination
                currentPage={resReportCurrentPage}
                totalPages={
                  resReportFilterData.length > 0
                    ? Math.ceil(resReportFilterData.length / totalPageItems)
                    : 1
                }
                onChange={(page) => {
                  paginateFunction(page, resReportFilterData);
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
                No data found !!!
              </Label>
            </div>
          )}
          {/* Body-Section Ends */}
        </div>
      )}
    </div>
  );
};

export default WRReviewReport;
