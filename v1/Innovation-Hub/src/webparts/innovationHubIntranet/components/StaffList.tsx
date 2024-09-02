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
  Rating,
  RatingSize,
  IIconProps,
} from "@fluentui/react";

import Service from "../components/Services";

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
  Role: string;
  UserName: string;
  UserEmail: string;
  UserID: number;
  ActiveStatus: string;
  State: string;
  ActualHours: number;
  NormalWeeklyWorkHours: number;
  PlannedHours: number;
  SetWeeklyProductionHours: number;
  Rating: number;
  OverallRating: number;
  OverallReview: string;
  RatingDetails: any;
}
interface IFilter {
  BA: string;
  Role: string;
  ActiveStatus: string;
  User: string;
}

interface IDropdowns {
  BA: [{ key: string; text: string }];
  Role: [{ key: string; text: string }];
  ActiveStatus: any;
}

let sortSlData: IData[] = [];
let sortSlFilterData: IData[] = [];

let globalMasterUserListData = [];
let globalPBData = [];
let globalAPBData = [];
let globalDRData = [];

const StaffList = (props: IProps): JSX.Element => {
  const sharepointWeb: any = Web(props.URL);
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  const allPeoples: any[] = props.peopleList;
  let CurrentPage: number = 1;
  let totalPageItems: number = 10;
  let thisWeek: number = moment().isoWeek();
  let thisYear: number = moment().year();

  const saveIcon: IIconProps = { iconName: "Save" };
  const editIcon: IIconProps = { iconName: "Edit" };
  const cancelIcon: IIconProps = { iconName: "Cancel" };

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

  const _SlColumns: any[] = props.isAdmin
    ? [
        {
          key: "Column1",
          name: "User",
          fieldName: "UserName",
          minWidth: 100,
          maxWidth: 250,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <div style={{ display: "flex" }}>
              <div
                style={{
                  marginTop: "-6px",
                }}
                title={item.UserName}
              >
                <Persona
                  size={PersonaSize.size32}
                  presence={PersonaPresence.none}
                  imageUrl={
                    "/_layouts/15/userphoto.aspx?size=S&username=" +
                    `${item.UserEmail}`
                  }
                />
              </div>
              <div>
                <span title={item.UserName} style={{ fontSize: "13px" }}>
                  {item.UserName}
                </span>
              </div>
            </div>
          ),
        },
        {
          key: "Column2",
          name: "Business area",
          fieldName: "BA",
          minWidth: 50,
          maxWidth: 110,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => {
            let BA = BAacronymsCollection.filter((ba) => {
              return ba.Name == item.BA;
            });
            return BA.length > 0 ? BA[0].ShortName : null;
          },
        },
        {
          key: "Column3",
          name: "Role",
          fieldName: "Role",
          minWidth: 100,
          maxWidth: 180,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              <TooltipHost
                id={item.ID}
                content={item.Role}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Role}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "Column4",
          name: (
            <TooltipHost content="Normal weekly work hours">
              <span style={{ cursor: "pointer" }}>
                Normal weekly work hours
              </span>
            </TooltipHost>
          ),
          fieldName: "NormalWeeklyWorkHours",
          minWidth: 60,
          maxWidth: 200,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) =>
            item.NormalWeeklyWorkHours.toString().match(/\./g)
              ? item.NormalWeeklyWorkHours.toFixed(2)
              : item.NormalWeeklyWorkHours,
        },
        {
          key: "Column5",
          name: (
            <TooltipHost content="Set weekly production hours">
              <span style={{ cursor: "pointer" }}>
                Set weekly production hours
              </span>
            </TooltipHost>
          ),
          fieldName: "SetWeeklyProductionHours",
          minWidth: 60,
          maxWidth: 200,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) =>
            item.SetWeeklyProductionHours.toString().match(/\./g)
              ? item.SetWeeklyProductionHours.toFixed(2)
              : item.SetWeeklyProductionHours,
        },
        {
          key: "Column6",
          name: "Planned hours",
          fieldName: "PlannedHours",
          minWidth: 100,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) =>
            item.PlannedHours.toString().match(/\./g)
              ? item.PlannedHours.toFixed(2)
              : item.PlannedHours,
        },
        {
          key: "Column7",
          name: "Actual hours",
          fieldName: "ActualHours",
          minWidth: 80,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) =>
            item.ActualHours.toString().match(/\./g)
              ? item.ActualHours.toFixed(2)
              : item.ActualHours,
        },
        // {
        //   key: "Column8",
        //   name: "Rating",
        //   fieldName: "OverallRating",
        //   minWidth: 105,
        //   maxWidth: 150,
        //   onColumnClick: (
        //     ev: React.MouseEvent<HTMLElement>,
        //     column: IColumn
        //   ) => {
        //     _onColumnClick(ev, column);
        //   },
        //   onRender: (item) => (
        //     <>
        //       <Rating
        //         max={4}
        //         rating={item.OverallRating}
        //         allowZeroStars
        //         styles={
        //           item.OverallRating >= 3
        //             ? {
        //                 ratingStarFront: {
        //                   color: "#00a300",
        //                   ".ms-Rating-button:hover .ms-RatingStar-front": {
        //                     color: "#00a300 !important",
        //                   },
        //                 },
        //                 ratingButton: {
        //                   cursor: "default !important",
        //                   ":hover .ms-RatingStar-front": {
        //                     color: "#00a300 !important",
        //                   },
        //                   ":hover ~ .ms-Rating-button .ms-RatingStar-front": {
        //                     color: "#00a300 !important",
        //                   },
        //                   ":hover .ms-RatingStar-back": {
        //                     color: "#605e5c !important",
        //                   },
        //                 },
        //                 ratingStarBack: {
        //                   color: "#605e5c !important",
        //                 },
        //               }
        //             : item.OverallRating >= 2
        //             ? {
        //                 ratingStarFront: { color: "#a3a300" },
        //                 ratingButton: {
        //                   cursor: "default !important",
        //                   ":hover .ms-RatingStar-front": {
        //                     color: "#a3a300 !important",
        //                   },
        //                   ":hover ~ .ms-Rating-button .ms-RatingStar-front": {
        //                     color: "#a3a300 !important",
        //                   },
        //                   ":hover .ms-RatingStar-back": {
        //                     color: "#605e5c !important",
        //                   },
        //                 },
        //                 ratingStarBack: {
        //                   color: "#605e5c !important",
        //                 },
        //               }
        //             : item.OverallRating >= 1
        //             ? {
        //                 ratingStarFront: { color: "#D18700" },
        //                 ratingButton: {
        //                   cursor: "default !important",
        //                   ":hover .ms-RatingStar-front": {
        //                     color: "#D18700 !important",
        //                   },
        //                   ":hover ~ .ms-Rating-button .ms-RatingStar-front": {
        //                     color: "#D18700 !important",
        //                   },
        //                   ":hover .ms-RatingStar-back": {
        //                     color: "#605e5c !important",
        //                   },
        //                 },
        //                 ratingStarBack: {
        //                   color: "#605e5c !important",
        //                 },
        //               }
        //             : item.OverallRating > 0
        //             ? {
        //                 ratingStarFront: { color: "#D10000" },
        //                 ratingButton: {
        //                   cursor: "default !important",
        //                   ":hover .ms-RatingStar-front": {
        //                     color: "#D10000 !important",
        //                   },
        //                   ":hover ~ .ms-Rating-button .ms-RatingStar-front": {
        //                     color: "#D10000 !important",
        //                   },
        //                   ":hover .ms-RatingStar-back": {
        //                     color: "#605e5c !important",
        //                   },
        //                 },
        //                 ratingStarBack: {
        //                   color: "#605e5c !important",
        //                 },
        //               }
        //             : null
        //         }
        //         disabled={false}
        //         style={{ width: 120 }}
        //         size={RatingSize.Large}
        //       />
        //       <Icon
        //         style={{
        //           color: "#2392b2",
        //         }}
        //         iconName="Info"
        //         className={SliconStyleClass.link}
        //         onClick={(_) => {
        //           let arrayRating = [
        //             {
        //               Rating: 4,
        //               Title: "Exceeds",
        //               Value: item.RatingDetails.Rating4 + " ratings",
        //               Percent:
        //                 item.RatingDetails.Rating4 > 0
        //                   ? (item.RatingDetails.Rating4 / item.OverallReview) *
        //                     100
        //                   : 0,
        //             },
        //             {
        //               Rating: 3,
        //               Title: "Achieved",
        //               Value: item.RatingDetails.Rating3 + " ratings",
        //               Percent:
        //                 item.RatingDetails.Rating3 > 0
        //                   ? (item.RatingDetails.Rating3 / item.OverallReview) *
        //                     100
        //                   : 0,
        //             },
        //             {
        //               Rating: 2,
        //               Title: "Developing",
        //               Value: item.RatingDetails.Rating2 + " ratings",
        //               Percent:
        //                 item.RatingDetails.Rating2 > 0
        //                   ? (item.RatingDetails.Rating2 / item.OverallReview) *
        //                     100
        //                   : 0,
        //             },
        //             {
        //               Rating: 1,
        //               Title: "Needs improvement",
        //               Value: item.RatingDetails.Rating1 + " ratings",
        //               Percent:
        //                 item.RatingDetails.Rating1 > 0
        //                   ? (item.RatingDetails.Rating1 / item.OverallReview) *
        //                     100
        //                   : 0,
        //             },
        //           ];
        //           setSlRatingPopup({
        //             condition: true,
        //             Rating: item.OverallRating,
        //             OverallReview: item.OverallReview,
        //             OverallRating: arrayRating,
        //           });
        //         }}
        //       />
        //     </>
        //   ),
        // },
      ]
    : [
        {
          key: "Column1",
          name: "User",
          fieldName: "UserName",
          minWidth: 100,
          maxWidth: 250,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <div style={{ display: "flex" }}>
              <div
                style={{
                  marginTop: "-6px",
                }}
                title={item.UserName}
              >
                <Persona
                  size={PersonaSize.size32}
                  presence={PersonaPresence.none}
                  imageUrl={
                    "/_layouts/15/userphoto.aspx?size=S&username=" +
                    `${item.UserEmail}`
                  }
                />
              </div>
              <div>
                <span title={item.UserName} style={{ fontSize: "13px" }}>
                  {item.UserName}
                </span>
              </div>
            </div>
          ),
        },
        {
          key: "Column2",
          name: "Business area",
          fieldName: "BA",
          minWidth: 50,
          maxWidth: 110,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => {
            let BA = BAacronymsCollection.filter((ba) => {
              return ba.Name == item.BA;
            });
            return BA.length > 0 ? BA[0].ShortName : null;
          },
        },
        {
          key: "Column3",
          name: "Role",
          fieldName: "Role",
          minWidth: 100,
          maxWidth: 180,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              <TooltipHost
                id={item.ID}
                content={item.Role}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Role}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "Column4",
          name: (
            <TooltipHost content="Normal weekly work hours">
              <span style={{ cursor: "pointer" }}>
                Normal weekly work hours
              </span>
            </TooltipHost>
          ),
          fieldName: "NormalWeeklyWorkHours",
          minWidth: 60,
          maxWidth: 200,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "Column5",
          name: (
            <TooltipHost content="Set weekly production hours">
              <span style={{ cursor: "pointer" }}>
                Set weekly production hours
              </span>
            </TooltipHost>
          ),
          fieldName: "SetWeeklyProductionHours",
          minWidth: 60,
          maxWidth: 200,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "Column6",
          name: "Planned hours",
          fieldName: "PlannedHours",
          minWidth: 100,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "Column7",
          name: "Actual hours",
          fieldName: "ActualHours",
          minWidth: 80,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
      ];
  const SlbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const SlbuttonStyleClass = mergeStyleSets({
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
      SlbuttonStyle,
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
      SlbuttonStyle,
    ],
  });
  const SlstatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    height: 10,
    width: 260,
  });
  const SlstatusStyleClass = mergeStyleSets({
    rating4: [
      {
        fontWeight: "600",
        backgroundColor: "#00a300",
      },
      SlstatusStyle,
    ],
    rating3: [
      {
        fontWeight: "600",
        backgroundColor: "#a3a300",
      },
      SlstatusStyle,
    ],
    rating2: [
      {
        fontWeight: "600",
        backgroundColor: "#D18700",
      },
      SlstatusStyle,
    ],
    rating1: [
      {
        fontWeight: "600",
        backgroundColor: "#D10000",
      },
      SlstatusStyle,
    ],
    default: [
      {
        fontWeight: "600",
        position: "relative",
        backgroundColor: "#edebe9",
        marginTop: 16,
      },
      SlstatusStyle,
    ],
    percentageText: [
      {
        position: "absolute !important",
        left: "50%",
        top: "50%",
        transform: "translate(-50%,-50%)",
        color: "#555",
      },
    ],
  });
  const SlFilterKeys: IFilter = {
    BA: "All",
    Role: "All",
    ActiveStatus: "Active",
    User: "",
  };
  const SllabelStyles = mergeStyleSets({
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
        marginRight: "10px",
      },
    ],
  });
  const SliconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const SliconStyleClass = mergeStyleSets({
    link: [{ color: "blue", margin: "0 0" }, SliconStyle],
    delete: [{ color: "red", margin: "0 7px" }, SliconStyle],
    edit: [{ color: "blue", margin: "0 7px" }, SliconStyle],
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
  const SlDropdownStyles: Partial<IDropdownStyles> = {
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
  const SlActiveDropdownStyles: Partial<IDropdownStyles> = {
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
  const SlSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 165,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
      marginTop: "3px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const SlActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 165,
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
  const SlLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const SlFilterOptns: IDropdowns = {
    BA: [{ key: "All", text: "All" }],
    Role: [{ key: "All", text: "All" }],
    ActiveStatus: [
      { key: "All", text: "All" },
      { key: "Active", text: "Active" },
      { key: "Inactive", text: "Inactive" },
    ],
  };

  // UseState
  const [SlReRender, setSlReRender] = useState<boolean>(false);
  const [SlMasterData, setSlMasterData] = useState<IData[]>([]);
  const [SlData, setSlData] = useState<IData[]>([]);
  const [SlFilterData, setSlFilterData] = useState<IData[]>([]);
  const [SlDisplayData, setSlDisplayData] = useState<IData[]>([]);
  const [SlLoader, setSlLoader] = useState<boolean>(true);
  const [SlFilter, setSlFilter] = useState<IFilter>(SlFilterKeys);
  const [SlFilterDrpDown, setSlFilterDrpDown] =
    useState<IDropdowns>(SlFilterOptns);
  const [SlColumns, setSlColumns] = useState(_SlColumns);
  const [SlcurrentPage, setSlCurrentPage] = useState<number>(CurrentPage);
  const [SlRatingPopup, setSlRatingPopup] = useState({
    condition: false,
    Rating: 0,
    OverallReview: 0,
    OverallRating: [],
  });

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _SlColumns;
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

    const newSlData = _copyAndSort(
      sortSlData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newSlFilterData = _copyAndSort(
      sortSlFilterData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setSlData([...newSlData]);
    setSlFilterData([...newSlFilterData]);
    paginateFunction(1, [...newSlFilterData]);
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

  //Functions for loading
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
    filterCondition: string
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
          getPagedValues(listName, filterCondition, data.NextHref);
        } else {
          listName == "ProductionBoard"
            ? getActivityProductionBoardData()
            : listName == "ActivityProductionBoard"
            ? getReviewLogData()
            : listName == "Review Log"
            ? dataManipulationFunction()
            : null;
        }
      })
      .catch((err: string) => {
        SlErrorFunction(err, `getThresholdData`);
      });
  };
  const getPagedValues = (
    listName: string,
    filterCondition: string,
    nextHref: string
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
          getPagedValues(listName, filterCondition, data.NextHref);
        } else {
          listName == "ProductionBoard"
            ? getActivityProductionBoardData()
            : listName == "ActivityProductionBoard"
            ? getReviewLogData()
            : listName == "Review Log"
            ? dataManipulationFunction()
            : null;
        }
      })
      .catch((err: string) => {
        SlErrorFunction(err, `getPagedValues`);
      });
  };
  const getMasterUserListData = (): void => {
    const sortFilterKeys = (a, b) => {
      if (a.Title < b.Title) {
        return -1;
      }
      if (a.Title > b.Title) {
        return 1;
      }
      return 0;
    };
    // setDBLoader("StartLoader");
    globalMasterUserListData = [];

    sharepointWeb.lists
      .getByTitle("Master User List")
      .items.select("*,User/EMail,User/Title")
      .expand("User")
      .top(5000)
      .get()
      .then((items) => {
        items = items.filter((user) => {
          return user.UserId;
        });
        globalMasterUserListData.push(...items);
        globalMasterUserListData.sort(sortFilterKeys);
        getProductionBoardData();
      })
      .catch((err: string) => {
        SlErrorFunction(err, "getMasterUserListData");
      });
  };
  const getProductionBoardData = (): void => {
    globalPBData = [];
    let queryArr = [
      `<Eq>
      <FieldRef Name='Week' />
      <Value Type='Number'>${thisWeek}</Value>
   </Eq>`,
      `<Eq>
      <FieldRef Name='Year' />
      <Value Type='Number'>${thisYear}</Value>
   </Eq>`,
    ];
    let productionBoardQuery = queryGenerator(queryArr);
    // console.log(productionBoardQuery);

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
      </ViewFields>
      <RowLimit Paged='TRUE'>5000</RowLimit>
    </View>`;

    getThresholdData("ProductionBoard", Filtercondition);
  };
  const getActivityProductionBoardData = (): void => {
    globalAPBData = [];
    let queryArr = [
      `<Eq>
      <FieldRef Name='Week' />
      <Value Type='Number'>${thisWeek}</Value>
   </Eq>`,
      `<Eq>
      <FieldRef Name='Year' />
      <Value Type='Number'>${thisYear}</Value>
   </Eq>`,
    ];
    let activityProductionBoardQuery = queryGenerator(queryArr);
    // console.log(activityProductionBoardQuery);

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
      </ViewFields>
      <RowLimit Paged='TRUE'>5000</RowLimit>
    </View>`;

    getThresholdData("ActivityProductionBoard", Filtercondition);
  };
  const getReviewLogData = (): void => {
    globalDRData = [];
    let dateOfaWeek = moment().isoWeek(thisWeek).year(thisYear);

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
        <FieldRef Name='auditFrom' />
        <FieldRef Name='auditTo' />
        <FieldRef Name='FromUser' />
        <FieldRef Name='ToUser' />
        <FieldRef Name='auditResponseType' />
        <FieldRef Name='auditLink' />
        <FieldRef Name='FeedbackRepeated' />
        <FieldRef Name='auditSent' />
        <FieldRef Name='Title' />
        <FieldRef Name='Rating' />
      </ViewFields>
      <RowLimit Paged='TRUE'>5000</RowLimit>
    </View>`;

    getThresholdData("Review Log", Filtercondition);
  };
  const dataManipulationFunction = (): void => {
    let _Sldata = [];
    globalMasterUserListData.forEach((arr, index) => {
      let PB_PH = totalHours(arr.User.EMail, "PlannedHours", globalPBData);
      let APB_PH = totalHours(arr.User.EMail, "PlannedHours", globalAPBData);
      let PB_AH = totalHours(arr.User.EMail, "ActualHours", globalPBData);
      let APB_AH = totalHours(arr.User.EMail, "ActualHours", globalAPBData);
      let OverallRating = grandTotal(
        arr.User.EMail,
        "Rating",
        globalDRData,
        0,
        false
      );
      let OverallReview = grandTotal(
        arr.User.EMail,
        "Rating",
        globalDRData,
        0,
        true
      );
      _Sldata.push({
        ID: index,
        UserID: arr.UserId ? arr.UserId : 0,
        UserName: arr.UserId ? arr.User.Title : "",
        UserEmail: arr.UserId ? arr.User.EMail : "",
        BA: arr.BusinessArea ? arr.BusinessArea : "",
        ActiveStatus: arr.Active ? "Active" : "Inactive",
        Role: arr.Position ? arr.Position : "",
        State: arr.State ? arr.State : "",
        NormalWeeklyWorkHours: arr.NormalWeeklyWorkHours
          ? arr.NormalWeeklyWorkHours
          : 0,
        SetWeeklyProductionHours: arr.SetWeeklyProductionHours
          ? arr.SetWeeklyProductionHours
          : 0,
        PlannedHours: PB_PH + APB_PH,
        ActualHours: PB_AH + APB_AH,
        OverallRating: OverallRating ? OverallRating / OverallReview : 0,
        OverallReview: OverallReview ? OverallReview : 0,
        RatingDetails: {
          Rating1: grandTotal(arr.User.EMail, "Rating", globalDRData, 1, true),
          Rating2: grandTotal(arr.User.EMail, "Rating", globalDRData, 2, true),
          Rating3: grandTotal(arr.User.EMail, "Rating", globalDRData, 3, true),
          Rating4: grandTotal(arr.User.EMail, "Rating", globalDRData, 4, true),
        },
      });
    });
    setSlData([..._Sldata]);
    sortSlData = _Sldata;
    setSlMasterData([..._Sldata]);
    reloadFilterDropdowns([..._Sldata]);

    let filterArr = _Sldata.filter((arr) => {
      return arr.ActiveStatus == "Active";
    });
    setSlFilterData([...filterArr]);
    sortSlFilterData = filterArr;
    paginateFunction(1, filterArr);
    setSlLoader(false);

    console.log(_Sldata, globalDRData);
  };
  const grandTotal = (userEmail, field, data, rating, lengthFlag) => {
    let sum = 0;
    data = data.filter((arr) => {
      return (
        (arr.FromUser ? arr.FromUser[0].email == userEmail : null) && arr.Rating
      );
    });
    if (rating != 0) {
      data = data.filter((arr) => {
        return arr.Rating ? arr.Rating == rating : null;
      });
    }
    data.forEach((arr) => {
      parseFloat(arr[field]) ? (sum += parseFloat(arr[field])) : 0;
    });
    if (lengthFlag) {
      return data.length;
    } else {
      return sum;
    }
  };
  const totalHours = (userEmail, field, data) => {
    let sum = 0;
    data = data.filter((arr) => {
      return arr.Developer ? arr.Developer[0].email == userEmail : null;
    });
    data.forEach((arr) => {
      parseFloat(arr[field]) ? (sum += parseFloat(arr[field])) : 0;
    });

    return sum;
  };

  //Getting data for loading
  const reloadFilterDropdowns = (data: IData[]): void => {
    let tempArrReload = data;

    tempArrReload.forEach((item) => {
      if (
        SlFilterOptns.BA.findIndex((BA) => {
          return BA.key == item.BA;
        }) == -1 &&
        item.BA
      ) {
        SlFilterOptns.BA.push({
          key: item.BA,
          text: item.BA,
        });
      }
      if (
        SlFilterOptns.Role.findIndex((Title) => {
          return Title.key == item.Role;
        }) == -1 &&
        item.Role
      ) {
        SlFilterOptns.Role.push({
          key: item.Role,
          text: item.Role,
        });
      }
    });
    setSlFilterDrpDown(SlFilterOptns);
  };
  const paginateFunction = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      CurrentPage = pagenumber;
      setSlDisplayData(paginatedItems);
      setSlCurrentPage(pagenumber);
    } else {
      setSlDisplayData([]);
      setSlCurrentPage(1);
    }
  };
  //Onchange Function
  const onChangeFilter = (key: string, option: string) => {
    let tempData: IData[] = SlData;
    let tempFilterKeys = SlFilter;
    tempFilterKeys[key] = option;
    if (tempFilterKeys.BA != "All") {
      tempData = tempData.filter((arr) => {
        return arr.BA == tempFilterKeys.BA;
      });
    }
    if (tempFilterKeys.Role != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Role == tempFilterKeys.Role;
      });
    }
    if (tempFilterKeys.ActiveStatus != "All") {
      tempData = tempData.filter((arr) => {
        return arr.ActiveStatus == tempFilterKeys.ActiveStatus;
      });
    }
    if (tempFilterKeys.User) {
      tempData = tempData.filter((arr) => {
        return arr.UserName.toLowerCase().includes(
          tempFilterKeys.User.toLowerCase()
        );
      });
    }
    setSlFilterData([...tempData]);
    sortSlFilterData = tempData;
    setSlFilter({ ...tempFilterKeys });
    paginateFunction(1, tempData);
  };
  const generateExcel = () => {
    let arrExport = SlFilterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "User", key: "User", width: 80 },
      { header: "Business area", key: "BA", width: 25 },
      { header: "Role", key: "Role", width: 25 },

      { header: "State", key: "State", width: 25 },
      {
        header: "NormalWeeklyWorkHours",
        key: "NormalWeeklyWorkHours",
        width: 25,
      },
      {
        header: "SetWeeklyProductionHours",
        key: "SetWeeklyProductionHours",
        width: 25,
      },
      { header: "PlannedHours", key: "PlannedHours", width: 20 },
      { header: "ActualHours", key: "ActualHours", width: 25 },
      // { header: "Rating", key: "Rating", width: 20 },
    ];
    arrExport.forEach((item) => {
      worksheet.addRow({
        BA: item.BA ? item.BA : "",
        Role: item.Role ? item.Role : "",
        User: item.UserName ? item.UserName : "",
        State: item.State ? item.State : "",
        NormalWeeklyWorkHours: item.NormalWeeklyWorkHours
          ? item.NormalWeeklyWorkHours
          : 0,
        SetWeeklyProductionHours: item.SetWeeklyProductionHours
          ? item.SetWeeklyProductionHours
          : 0,
        PlannedHours: item.PlannedHours ? item.PlannedHours : 0,
        ActualHours: item.ActualHours ? item.ActualHours : 0,
        // Rating: item.Rating ? item.Rating : 0,
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
          `StaffList-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };
  const SlErrorFunction = (error: any, functionName: string) => {
    console.log(error, functionName);

    let response = {
      ComponentName: "Staff list",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        ErrorPopup();
      }
    );
  };
  const AddSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Staff list is successfully submitted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  //Use Effect
  useEffect(() => {
    setSlLoader(true);
    getMasterUserListData();
  }, [SlReRender]);

  return (
    <div style={{ padding: "5px 15px" }}>
      {SlLoader ? <CustomLoader /> : null}
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
            marginBottom: 15,
            color: "#2392b2",
          }}
        >
          <div className={styles.dpTitle}>
            <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
              Staff List
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
          }}
        >
          <div className={styles.ddSection}>
            <div>
              <Label styles={SlLabelStyles}>User</Label>
              <SearchBox
                styles={
                  SlFilter.User ? SlActiveSearchBoxStyles : SlSearchBoxStyles
                }
                value={SlFilter.User}
                onChange={(e, value): void => {
                  onChangeFilter("User", value);
                }}
              />
            </div>
            <div>
              <Label styles={SlLabelStyles}>Business area</Label>
              <Dropdown
                placeholder="Select an option"
                options={SlFilterDrpDown.BA}
                selectedKey={SlFilter.BA}
                styles={
                  SlFilter.BA == "All"
                    ? SlDropdownStyles
                    : SlActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("BA", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={SlLabelStyles}>Role</Label>
              <Dropdown
                placeholder="Select an option"
                options={SlFilterDrpDown.Role}
                selectedKey={SlFilter.Role}
                styles={
                  SlFilter.Role == "All"
                    ? SlDropdownStyles
                    : SlActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("Role", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={SlLabelStyles}>User state</Label>
              <Dropdown
                placeholder="Select an option"
                options={SlFilterDrpDown.ActiveStatus}
                selectedKey={SlFilter.ActiveStatus}
                styles={
                  SlFilter.ActiveStatus == "All"
                    ? SlDropdownStyles
                    : SlActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("ActiveStatus", option["key"]);
                }}
              />
            </div>
            <div>
              <div>
                <Icon
                  iconName="Refresh"
                  title="Click to reset"
                  className={SliconStyleClass.refresh}
                  onClick={() => {
                    setSlData([...SlMasterData]);
                    sortSlData = SlMasterData;

                    let filterArr = SlMasterData.filter((arr) => {
                      return arr.ActiveStatus == "Active";
                    });
                    setSlFilterData([...filterArr]);
                    sortSlFilterData = filterArr;

                    paginateFunction(1, SlMasterData);
                    setSlFilter({ ...SlFilterKeys });
                    setSlColumns(_SlColumns);
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
            }}
          >
            <Label className={SllabelStyles.NORLabel}>
              Number of records:{" "}
              <b style={{ color: "#038387" }}>{SlFilterData.length}</b>
            </Label>
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
                }}
              >
                <Icon
                  style={{
                    color: "#1D6F42",
                  }}
                  iconName="ExcelDocument"
                  className={SliconStyleClass.export}
                />
                Export as XLS
              </Label>
            </div>
          </div>
        </div>
      </div>
      <div>
        <DetailsList
          items={SlDisplayData}
          columns={SlColumns}
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
      {SlFilterData.length > 0 ? (
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            margin: "10px 0",
          }}
        >
          <Pagination
            currentPage={SlcurrentPage}
            totalPages={
              SlFilterData.length > 0
                ? Math.ceil(SlFilterData.length / totalPageItems)
                : 1
            }
            onChange={(page) => {
              paginateFunction(page, SlFilterData);
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
      {SlRatingPopup.condition ? (
        <Modal
          isOpen={SlRatingPopup.condition}
          isBlocking={false}
          styles={{
            root: {},
            scrollableContent: {
              width: "530px",
            },
          }}
        >
          <div style={{ padding: "30px 20px" }}>
            <Label
              style={{
                textAlign: "center",
                color: "#2392b2",
                fontSize: 20,
                fontWeight: 600,
              }}
            >
              Overall Rating
            </Label>
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
                  setSlRatingPopup({
                    condition: false,
                    Rating: 0,
                    OverallReview: 0,
                    OverallRating: [],
                  });
                }}
              />
            </div>
            <Label
              style={{
                textAlign: "center",
                fontSize: 40,
                fontWeight: 600,
              }}
            >
              {SlRatingPopup.Rating.toString().match(/\./g)
                ? SlRatingPopup.Rating.toFixed(2) + " / 4"
                : SlRatingPopup.Rating + " / 4"}
            </Label>
            <div
              style={{
                display: "flex",
                justifyContent: "center",
              }}
            >
              <Rating
                max={4}
                rating={SlRatingPopup.Rating}
                allowZeroStars
                styles={
                  SlRatingPopup.Rating >= 3
                    ? {
                        ratingStarFront: { color: "#00a300" },
                        ratingButton: {
                          ":hover .ms-RatingStar-front": {
                            color: "#00a300 !important",
                          },
                          ":hover ~ .ms-Rating-button .ms-RatingStar-front": {
                            color: "#00a300 !important",
                          },
                          ":hover .ms-RatingStar-back": {
                            color: "#605e5c !important",
                          },
                        },
                        ratingStarBack: {
                          color: "#605e5c !important",
                        },
                      }
                    : SlRatingPopup.Rating >= 2
                    ? {
                        ratingStarFront: { color: "#a3a300" },
                        ratingButton: {
                          ":hover .ms-RatingStar-front": {
                            color: "#a3a300 !important",
                          },
                          ":hover ~ .ms-Rating-button .ms-RatingStar-front": {
                            color: "#a3a300 !important",
                          },
                          ":hover .ms-RatingStar-back": {
                            color: "#605e5c !important",
                          },
                        },
                        ratingStarBack: {
                          color: "#605e5c !important",
                        },
                      }
                    : SlRatingPopup.Rating >= 1
                    ? {
                        ratingStarFront: { color: "#D18700" },
                        ratingButton: {
                          ":hover .ms-RatingStar-front": {
                            color: "#D18700 !important",
                          },
                          ":hover ~ .ms-Rating-button .ms-RatingStar-front": {
                            color: "#D18700 !important",
                          },
                          ":hover .ms-RatingStar-back": {
                            color: "#605e5c !important",
                          },
                        },
                        ratingStarBack: {
                          color: "#605e5c !important",
                        },
                      }
                    : SlRatingPopup.Rating > 0
                    ? {
                        ratingStarFront: { color: "#D10000" },
                        ratingButton: {
                          ":hover .ms-RatingStar-front": {
                            color: "#D10000 !important",
                          },
                          ":hover ~ .ms-Rating-button .ms-RatingStar-front": {
                            color: "#D10000 !important",
                          },
                          ":hover .ms-RatingStar-back": {
                            color: "#605e5c !important",
                          },
                        },
                        ratingStarBack: {
                          color: "#605e5c !important",
                        },
                      }
                    : null
                }
                disabled={false}
                // style={{ width: 120 }}
                size={RatingSize.Large}
              />
            </div>
            <Label
              style={{
                textAlign: "center",
                fontSize: 10,
                fontWeight: 400,
                marginBottom: "30px",
              }}
            >
              {"Based on " + SlRatingPopup.OverallReview + " reviews"}
            </Label>
            {SlRatingPopup.OverallRating.map((OAR) => {
              return (
                <div
                  style={{
                    display: "flex",
                    justifyContent: "flex-start",
                  }}
                >
                  <Label
                    style={{
                      padding: "10px 20px",
                      fontSize: 13,
                      width: "140px",
                    }}
                  >
                    {OAR.Title}
                  </Label>
                  <div className={SlstatusStyleClass.default}>
                    <div
                      style={{
                        width: `${OAR.Percent}%`,
                      }}
                      // style={{ width: "100%" }}
                      className={
                        OAR.Rating == 4
                          ? SlstatusStyleClass.rating4
                          : OAR.Rating == 3
                          ? SlstatusStyleClass.rating3
                          : OAR.Rating == 2
                          ? SlstatusStyleClass.rating2
                          : SlstatusStyleClass.rating1
                      }
                    >
                      {/* {item.Completion} */}
                    </div>
                  </div>
                  <Label
                    style={{
                      padding: "10px 20px",
                      fontSize: 13,
                      width: "110px",
                      textAlign: "right",
                    }}
                  >
                    {OAR.Value}
                  </Label>
                </div>
              );
            })}
          </div>
        </Modal>
      ) : null}
    </div>
  );
};
export default StaffList;
