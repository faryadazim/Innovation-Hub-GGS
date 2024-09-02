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
  Rating,
  RatingSize,
  IIconProps,
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
  Role: string;
  User: string;
  UserId: number;
  Allocation: number;
  ActualHrsPerWeek: number;
  HrsPerWeek: number;
  ActualHrsPerDay: number;
  HrsPerDay: number;
  Rating: number;
  OverallRating: String;
}
interface IFilter {
  BA: string;
  Role: string;
}

interface IDropdowns {
  BA: [{ key: string; text: string }];
  Role: [{ key: string; text: string }];
}

let sortSlData: IData[] = [];
let sortSlFilterData: IData[] = [];
let sortSlUpdate: boolean = false;

const StaffList = (props: IProps): JSX.Element => {
  const sharepointWeb: any = Web(props.URL);
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  const allPeoples: any[] = props.peopleList;
  let CurrentPage: number = 1;
  let totalPageItems: number = 10;

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

  const _SlColumns: IColumn[] = props.isAdmin
    ? [
        {
          key: "Column1",
          name: "Name",
          fieldName: "User",
          minWidth: 100,
          maxWidth: 250,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
          onRender: (item) => (
            <div style={{ display: "flex" }}>
              <div
                style={{
                  marginTop: "-6px",
                }}
                title={item.User}
              >
                <Persona
                  size={PersonaSize.size32}
                  presence={PersonaPresence.none}
                  imageUrl={
                    "/_layouts/15/userphoto.aspx?size=S&username=" +
                    `${
                      allPeoples.filter((ap) => {
                        return ap.ID == item.UserId;
                      }).length > 0
                        ? allPeoples.filter((ap) => {
                            return ap.ID == item.UserId;
                          })[0].secondaryText
                        : null
                    }`
                  }
                />
              </div>
              <div>
                <span title={item.Provider} style={{ fontSize: "13px" }}>
                  {item.User}
                </span>
              </div>
            </div>
          ),
        },
        {
          key: "Column2",
          name: "Business area",
          fieldName: "BA",
          minWidth: 110,
          maxWidth: 110,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
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
          maxWidth: 250,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
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
          name: "Allocation%",
          fieldName: "Allocation",
          minWidth: 100,
          maxWidth: 100,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
          onRender: (item, Index) =>
            sortSlUpdate ? (
              <>
                <TextField
                  styles={{
                    root: {
                      selectors: {
                        ".ms-TextField-fieldGroup": {
                          borderRadius: 4,
                          border: "1px solid",
                          height: 28,
                          width: 70,
                          input: {
                            borderRadius: 4,
                          },
                        },
                      },
                    },
                  }}
                  data-id={item.ID}
                  disabled={false}
                  value={item.Allocation}
                  onChange={(e: any) => {
                    parseInt(e.target.value)
                      ? SlOnchangeItems(item.ID, "Allocation", e.target.value)
                      : SlOnchangeItems(item.ID, "Allocation", null);
                  }}
                />
              </>
            ) : (
              item.Allocation
            ),
        },
        {
          key: "Column5",
          name: "Actual Hrs/Week",
          fieldName: "ActualHrsPerWeek",
          minWidth: 130,
          maxWidth: 130,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
        },
        {
          key: "Column6",
          name: "Hrs/Week",
          fieldName: "HrsPerWeek",
          minWidth: 80,
          maxWidth: 80,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
        },
        {
          key: "Column7",
          name: "Actual Hrs/Day",
          fieldName: "ActualHrsPerDay",
          minWidth: 130,
          maxWidth: 130,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
        },
        {
          key: "Column8",
          name: "Hrs/Day",
          fieldName: "HrsPerDay",
          minWidth: 80,
          maxWidth: 80,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
        },
        {
          key: "Column9",
          name: "Rating",
          fieldName: "Rating",
          minWidth: 130,
          maxWidth: 130,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
          onRender: (item) => (
            <>
              <Rating
                max={4}
                rating={item.Rating}
                allowZeroStars
                styles={
                  item.Rating == 4
                    ? {
                        ratingStarFront: {
                          color: "#00a300",
                          ".ms-Rating-button:hover .ms-RatingStar-front": {
                            color: "#00a300 !important",
                          },
                        },
                        ratingButton: {
                          cursor: "default !important",
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
                    : item.Rating == 3
                    ? {
                        ratingStarFront: { color: "#a3a300" },
                        ratingButton: {
                          cursor: "default !important",
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
                    : item.Rating == 2
                    ? {
                        ratingStarFront: { color: "#D18700" },
                        ratingButton: {
                          cursor: "default !important",
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
                    : {
                        ratingStarFront: { color: "#D10000" },
                        ratingButton: {
                          cursor: "default !important",
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
                }
                disabled={false}
                style={{ width: 120 }}
                size={RatingSize.Large}
                // onClick={(_) => {
                //   let arrRating = item.OverallRating.split(";");
                //   let arrayRating = [
                //     {
                //       Rating: 4,
                //       Title: "Exceeds",
                //       Value: arrRating[0] + " ratings",
                //       Percent:
                //         arrRating[0] > 0 ? (arrRating[0] / arrRating[4]) * 100 : 0,
                //     },
                //     {
                //       Rating: 3,
                //       Title: "Achieved",
                //       Value: arrRating[1] + " ratings",
                //       Percent:
                //         arrRating[1] > 0 ? (arrRating[1] / arrRating[4]) * 100 : 0,
                //     },
                //     {
                //       Rating: 2,
                //       Title: "Developing",
                //       Value: arrRating[2] + " ratings",
                //       Percent:
                //         arrRating[2] > 0 ? (arrRating[2] / arrRating[4]) * 100 : 0,
                //     },
                //     {
                //       Rating: 1,
                //       Title: "Needs improvement",
                //       Value: arrRating[3] + " ratings",
                //       Percent:
                //         arrRating[3] > 0 ? (arrRating[3] / arrRating[4]) * 100 : 0,
                //     },
                //   ];
                //   setSlRatingPopup({
                //     condition: true,
                //     Rating: item.Rating,
                //     OverallReview: arrRating[4],
                //     OverallRating: arrayRating,
                //   });
                // }}
              />
              <Icon
                style={{
                  color: "#2392b2",
                }}
                iconName="Info"
                className={SliconStyleClass.link}
                onClick={(_) => {
                  let arrRating = item.OverallRating.split(";");
                  let arrayRating = [
                    {
                      Rating: 4,
                      Title: "Exceeds",
                      Value: arrRating[0] + " ratings",
                      Percent:
                        arrRating[0] > 0
                          ? (arrRating[0] / arrRating[4]) * 100
                          : 0,
                    },
                    {
                      Rating: 3,
                      Title: "Achieved",
                      Value: arrRating[1] + " ratings",
                      Percent:
                        arrRating[1] > 0
                          ? (arrRating[1] / arrRating[4]) * 100
                          : 0,
                    },
                    {
                      Rating: 2,
                      Title: "Developing",
                      Value: arrRating[2] + " ratings",
                      Percent:
                        arrRating[2] > 0
                          ? (arrRating[2] / arrRating[4]) * 100
                          : 0,
                    },
                    {
                      Rating: 1,
                      Title: "Needs improvement",
                      Value: arrRating[3] + " ratings",
                      Percent:
                        arrRating[3] > 0
                          ? (arrRating[3] / arrRating[4]) * 100
                          : 0,
                    },
                  ];
                  setSlRatingPopup({
                    condition: true,
                    Rating: item.Rating,
                    OverallReview: arrRating[4],
                    OverallRating: arrayRating,
                  });
                }}
              />
            </>
          ),
        },
      ]
    : [
        {
          key: "Column1",
          name: "Name",
          fieldName: "User",
          minWidth: 100,
          maxWidth: 250,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
          onRender: (item) => (
            <div style={{ display: "flex" }}>
              <div
                style={{
                  marginTop: "-6px",
                }}
                title={item.User}
              >
                <Persona
                  size={PersonaSize.size32}
                  presence={PersonaPresence.none}
                  imageUrl={
                    "/_layouts/15/userphoto.aspx?size=S&username=" +
                    `${
                      allPeoples.filter((ap) => {
                        return ap.ID == item.UserId;
                      }).length > 0
                        ? allPeoples.filter((ap) => {
                            return ap.ID == item.UserId;
                          })[0].secondaryText
                        : null
                    }`
                  }
                />
              </div>
              <div>
                <span title={item.Provider} style={{ fontSize: "13px" }}>
                  {item.User}
                </span>
              </div>
            </div>
          ),
        },
        {
          key: "Column2",
          name: "Business area",
          fieldName: "BA",
          minWidth: 110,
          maxWidth: 130,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
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
          maxWidth: 250,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
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
          name: "Allocation%",
          fieldName: "Allocation",
          minWidth: 100,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
          onRender: (item, Index) =>
            sortSlUpdate ? (
              <>
                <TextField
                  styles={{
                    root: {
                      selectors: {
                        ".ms-TextField-fieldGroup": {
                          borderRadius: 4,
                          border: "1px solid",
                          height: 28,
                          width: 70,
                          input: {
                            borderRadius: 4,
                          },
                        },
                      },
                    },
                  }}
                  data-id={item.ID}
                  disabled={false}
                  value={item.Allocation}
                  onChange={(e: any) => {
                    parseInt(e.target.value)
                      ? SlOnchangeItems(item.ID, "Allocation", e.target.value)
                      : SlOnchangeItems(item.ID, "Allocation", null);
                  }}
                />
              </>
            ) : (
              item.Allocation
            ),
        },
        {
          key: "Column5",
          name: "Actual Hrs/Week",
          fieldName: "ActualHrsPerWeek",
          minWidth: 130,
          maxWidth: 150,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
        },
        {
          key: "Column6",
          name: "Hrs/Week",
          fieldName: "HrsPerWeek",
          minWidth: 80,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
        },
        {
          key: "Column7",
          name: "Actual Hrs/Day",
          fieldName: "ActualHrsPerDay",
          minWidth: 130,
          maxWidth: 150,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
          },
        },
        {
          key: "Column8",
          name: "Hrs/Day",
          fieldName: "HrsPerDay",
          minWidth: 80,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            !sortSlUpdate ? _onColumnClick(ev, column) : null;
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
      fontSize: 10,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 10,
    },
    caretDown: { fontSize: 14, color: "#000" },
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
  const [SlUpdate, setSlUpdate] = useState(false);

  //FUnctions
  const getPBDRList = () => {
    let PBDRData = [];

    // Review Log
    sharepointWeb.lists
      .getByTitle("Review Log")
      .items.top(5000)
      .orderBy("auditSent", false)
      .get()
      .then((itemsPBDR) => {
        // ProductionBoard
        sharepointWeb.lists
          .getByTitle("ProductionBoard")
          .items.filter(
            "Week eq '" +
              moment().isoWeek() +
              "' and Year eq '" +
              moment().year() +
              "' and ActualHours ne 0 and ActualHours ne null"
          )
          .top(5000)
          .orderBy("Modified", false)
          .get()
          .then((itemsPB) => {
            itemsPB.forEach((itemPB) => {
              let varRating = itemsPBDR.filter((arr) => {
                return (
                  arr.ProductionBoardID == itemPB.ID &&
                  arr.DRPageName == "Annual Plan" &&
                  arr.Rating != 0
                );
              });
              varRating.forEach((arr) => {
                PBDRData.push({
                  DeveloperId: itemPB.DeveloperId,
                  ActualHours: 0,
                  Rating: arr.Rating,
                });
              });
              PBDRData.push({
                DeveloperId: itemPB.DeveloperId,
                ActualHours: itemPB.ActualHours,
                Rating: 0,
              });
            });

            // ActivityProductionBoard
            sharepointWeb.lists
              .getByTitle("ActivityProductionBoard")
              .items.filter(
                "Week eq '" +
                  moment().isoWeek() +
                  "' and Year eq '" +
                  moment().year() +
                  "' and ActualHours ne 0 and ActualHours ne null"
              )
              .top(5000)
              .orderBy("Modified", false)
              .get()
              .then((itemsAPB) => {
                itemsAPB.forEach((itemAPB) => {
                  let varRating = itemsPBDR.filter((arr) => {
                    return (
                      arr.ProductionBoardID == itemAPB.ID &&
                      arr.DRPageName == "Activity Plan" &&
                      arr.Rating != 0
                    );
                  });

                  varRating.forEach((arr) => {
                    PBDRData.push({
                      DeveloperId: itemAPB.DeveloperId,
                      ActualHours: 0,
                      Rating: arr.Rating,
                    });
                  });

                  PBDRData.push({
                    DeveloperId: itemAPB.DeveloperId,
                    ActualHours: itemAPB.ActualHours,
                    Rating: 0,
                  });
                });
                getStaffList(PBDRData);
              })
              .catch((error) => {
                SlErrorFunction(error, "getAPBList");
              });
          })
          .catch((error) => {
            SlErrorFunction(error, "getPBList");
          });
      })
      .catch((error) => {
        SlErrorFunction(error, "getDRList");
      });
  };
  const getStaffList = (PBDRData) => {
    let _Sldata: IData[] = [];
    sharepointWeb.lists
      .getByTitle("StaffList")
      .items.select("*", "User/Title", "User/Id", "User/EMail")
      .expand("User")
      .top(5000)
      .orderBy("User/EMail", true)
      .get()
      .then(async (items) => {
        items.forEach((item) => {
          let ActualHrsPerWeek = 0;
          let Rating = 0;
          let RatingCount = 0;
          let curUser = PBDRData.filter((arr) => {
            return arr.DeveloperId == item.UserId;
          });
          curUser.forEach((arr) => {
            ActualHrsPerWeek += arr.ActualHours;
            if (arr.Rating != 0) {
              Rating += arr.Rating;
              RatingCount++;
            }
          });
          let Rating1 = PBDRData.filter((arr) => {
            return arr.DeveloperId == item.UserId && arr.Rating == 1;
          });
          let Rating2 = PBDRData.filter((arr) => {
            return arr.DeveloperId == item.UserId && arr.Rating == 2;
          });
          let Rating3 = PBDRData.filter((arr) => {
            return arr.DeveloperId == item.UserId && arr.Rating == 3;
          });
          let Rating4 = PBDRData.filter((arr) => {
            return arr.DeveloperId == item.UserId && arr.Rating == 4;
          });
          let OverAllRating =
            Rating1.length +
            ";" +
            Rating2.length +
            ";" +
            Rating3.length +
            ";" +
            Rating4.length +
            ";" +
            (Rating1.length + Rating2.length + Rating3.length + Rating4.length);

          _Sldata.push({
            ID: item.ID,
            BA: item.BA,
            Role: item.Title,
            User: item.User ? item.User.Title : "",
            UserId: item.UserId,
            Allocation: item.Allocation ? item.Allocation : 0,
            HrsPerWeek: item.HrsPerWeek,
            HrsPerDay: item.HrsPerDay,

            // ActualHrsPerWeek: item.ActualHrsPerWeek,
            // ActualHrsPerDay: item.ActualHrsPerDay,
            // Rating: item.Rating,
            // OverallRating: item.OverallRating,
            ActualHrsPerWeek: ActualHrsPerWeek,
            ActualHrsPerDay: ActualHrsPerWeek ? ActualHrsPerWeek / 5 : 0,
            Rating: Rating ? Rating / RatingCount : 0,
            OverallRating: OverAllRating,
          });
        });
        console.log(_Sldata);
        setSlFilterData([..._Sldata]);
        sortSlFilterData = _Sldata;
        setSlData([..._Sldata]);
        sortSlData = _Sldata;
        setSlMasterData([..._Sldata]);
        reloadFilterDropdowns([..._Sldata]);
        paginateFunction(1, _Sldata);
        setSlLoader(false);
      })
      .catch((error) => {
        SlErrorFunction(error, "getStaffList");
      });
  };
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
  const saveSlData = () => {
    setSlLoader(true);
    let successCount = 0;
    SlData.forEach((sl, Index: number) => {
      let requestdata = {
        Allocation: sl.Allocation ? sl.Allocation : 0,
        HrsPerWeek: sl.HrsPerWeek ? sl.HrsPerWeek : 0,
        HrsPerDay: sl.HrsPerDay ? sl.HrsPerDay : 0,
      };

      sharepointWeb.lists
        .getByTitle("StaffList")
        .items.getById(sl.ID)
        .update(requestdata)
        .then((e) => {
          successCount++;

          if (SlData.length == successCount) {
            setSlMasterData([...SlData]);
            setSlUpdate(false);
            sortSlUpdate = false;
            AddSuccessPopup();
            sortSlData = SlData;
            setSlLoader(false);
          }
        })
        .catch((error) => {
          SlErrorFunction(error, "UpdateList");
        });
    });
  };
  const generateExcel = () => {
    let arrExport = SlFilterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Business area", key: "BA", width: 25 },
      { header: "Role", key: "Role", width: 25 },
      { header: "User", key: "User", width: 25 },
      { header: "Allocation%", key: "Allocation", width: 25 },
      { header: "ActualHrsPerWeek", key: "ActualHrsPerWeek", width: 25 },
      { header: "HrsPerWeek", key: "HrsPerWeek", width: 25 },
      { header: "ActualHrsPerDay", key: "ActualHrsPerDay", width: 20 },
      { header: "HrsPerDay", key: "HrsPerDay", width: 25 },
      { header: "Rating", key: "Rating", width: 20 },
    ];
    arrExport.forEach((item) => {
      worksheet.addRow({
        BA: item.BA ? item.BA : "",
        Role: item.Role ? item.Role : "",
        User: item.User ? item.User : "",
        Allocation: item.Allocation ? item.Allocation : 0,
        ActualHrsPerWeek: item.ActualHrsPerWeek ? item.ActualHrsPerWeek : "",
        HrsPerWeek: item.HrsPerWeek ? item.HrsPerWeek : "",
        ActualHrsPerDay: item.ActualHrsPerDay ? item.ActualHrsPerDay : "",
        HrsPerDay: item.HrsPerDay ? item.HrsPerDay : "",
        Rating: item.Rating ? item.Rating : "",
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
          `StaffList-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };
  const SlErrorFunction = (error: any, funName: string) => {
    console.log(error, funName);
    ErrorPopup();
  };
  const AddSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Staff list is successfully submitted !!!")
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
    setSlFilterData([...tempData]);
    sortSlFilterData = tempData;
    setSlFilter({ ...tempFilterKeys });
    paginateFunction(1, tempData);
  };
  const SlOnchangeItems = (RefId, key, value) => {
    let Index = SlData.findIndex((obj) => obj.ID == RefId);
    let disIndex = SlDisplayData.findIndex((obj) => obj.ID == RefId);
    let SLBeforeData = SlData[Index];

    let varHrsPerWeek = (key = "Allocation"
      ? (value / 100) * 37.5
      : (SLBeforeData.Allocation / 100) * 37.5);
    let HrsPerDay: number = varHrsPerWeek / 5;
    let SlOnchangeData = [
      {
        ID: SLBeforeData.ID,
        BA: SLBeforeData.BA,
        Role: SLBeforeData.Role,
        User: SLBeforeData.User,
        UserId: SLBeforeData.UserId,
        Allocation: (key = "Allocation" ? value : SLBeforeData.Allocation),
        ActualHrsPerWeek: SLBeforeData.ActualHrsPerWeek,
        HrsPerWeek: varHrsPerWeek,
        ActualHrsPerDay: SLBeforeData.ActualHrsPerDay,
        HrsPerDay: HrsPerDay,
        Rating: SLBeforeData.Rating,
        OverallRating: SLBeforeData.OverallRating,
      },
    ];

    SlData[Index] = SlOnchangeData[0];
    SlDisplayData[disIndex] = SlOnchangeData[0];
    reloadFilterDropdowns(SlData);
    setSlData([...SlData]);
    sortSlData = SlData;
  };

  //Use Effect
  useEffect(() => {
    getPBDRList();
    //getStaffList();
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
                className={SliconStyleClass.export}
              />
              Export as XLS
            </Label>
            <div>
              {SlUpdate ? (
                <div>
                  <PrimaryButton
                    iconProps={cancelIcon}
                    text="Cancel"
                    className={SlbuttonStyleClass.buttonPrimary}
                    onClick={(_) => {
                      setSlUpdate(false);
                      sortSlUpdate = false;

                      setSlFilterData([...SlMasterData]);
                      setSlData([...SlMasterData]);
                      sortSlFilterData = SlMasterData;
                      sortSlData = SlMasterData;
                      paginateFunction(1, SlMasterData);
                      setSlFilter({ ...SlFilterKeys });
                      setSlColumns(_SlColumns);
                    }}
                  />
                  <PrimaryButton
                    iconProps={saveIcon}
                    text="Save"
                    id="pbBtnSave"
                    className={SlbuttonStyleClass.buttonSecondary}
                    onClick={(_) => {
                      saveSlData();
                    }}
                  />
                </div>
              ) : (
                <div>
                  <PrimaryButton
                    iconProps={editIcon}
                    text="Edit"
                    className={SlbuttonStyleClass.buttonPrimary}
                    onClick={() => {
                      setSlUpdate(true);
                      sortSlUpdate = true;

                      setSlColumns(_SlColumns);
                      setSlData([...SlMasterData]);
                      // setSlFilterData(sortSlFilterData);
                      // paginateFunction(1, sortSlFilterData);
                      sortSlData = SlMasterData;

                      let tempData: IData[] = SlMasterData;
                      let tempFilterKeys = SlFilter;
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
                      setSlFilterData([...tempData]);
                      sortSlFilterData = tempData;
                      paginateFunction(1, tempData);
                    }}
                  />
                  <PrimaryButton
                    iconProps={saveIcon}
                    text="Save"
                    disabled={true}
                    onClick={(_) => {}}
                  />
                </div>
              )}
            </div>
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
              <div>
                <Icon
                  iconName="Refresh"
                  title="Click to reset"
                  className={SliconStyleClass.refresh}
                  onClick={() => {
                    setSlFilterData([...SlMasterData]);
                    setSlData([...SlMasterData]);
                    sortSlFilterData = SlMasterData;
                    sortSlData = SlMasterData;
                    paginateFunction(1, SlMasterData);
                    setSlFilter({ ...SlFilterKeys });
                    setSlColumns(_SlColumns);

                    setSlUpdate(false);
                    sortSlUpdate = false;
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
            <Label className={SllabelStyles.NORLabel}>
              Number of records:{" "}
              <b style={{ color: "#038387" }}>{SlFilterData.length}</b>
            </Label>
          </div>
        </div>
      </div>
      <div>
        <DetailsList
          items={SlDisplayData}
          columns={sortSlUpdate ? _SlColumns : SlColumns}
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
              {SlRatingPopup.Rating % 1 == 0 ||
              SlRatingPopup.Rating.toString().match(/\./g)
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
                  SlRatingPopup.Rating == 4
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
                    : SlRatingPopup.Rating == 3
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
                    : SlRatingPopup.Rating == 2
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
                    : {
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
