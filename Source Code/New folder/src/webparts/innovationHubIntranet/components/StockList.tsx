import * as React from "react";
import { useState, useEffect } from "react";
import { IWeb, Web } from "@pnp/sp/webs";
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
  TextField,
  ITextFieldStyles,
  NormalPeoplePicker,
  SearchBox,
  ISearchBoxStyles,
  Dropdown,
  IDropdownStyles,
  Modal,
  IColumn,
  Spinner,
  TooltipHost,
  TooltipOverflowMode,
  Persona,
  PersonaPresence,
  PersonaSize,
  Pivot,
  PivotItem,
  IPivotStyles,
} from "@fluentui/react";

import "../ExternalRef/styleSheets/Styles.css";
import Pagination from "office-ui-fabric-react-pagination";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { IPeoplelist, IDropdownOption } from "./IInnovationHubIntranetProps";
import * as moment from "moment";

import Service from "../components/Services";

interface IProps {
  context: any;
  spcontext: any;
  graphContent: any;
  URL: string;
  handleclick: any;
  pageType: string;
  peopleList: IPeoplelist[];
  isAdmin: boolean;
  WeblistURL: string;
}

interface IData {
  ID: number;
  BA: string;
  Product: string;
  Project: string;
  Subject: string;
  Stocks: string;
  Location: any;
  // School: string;
  Availability: any;
  Request: any;
}

interface IRequestData {
  ID: number;
  requestFromName: string;
  requestFromId: number;
  requestFromEmail: string;
  requestQuantity: number;
  requestStatus: string;
  // stockName: string;
  stockID: number;
  createdDate: Date;
  comments: string;
  school: string;
}

interface ICurrentuser {
  Name: string;
  Email: string;
  Id: number;
}

interface ISolDrpdwn {
  baOptns: IDropdownOption[];
  ProductOptns: IDropdownOption[];
  ProjectOptns: IDropdownOption[];
  SubjectOptns: IDropdownOption[];
  schoolOptns: IDropdownOption[];
  LocationOptns: IDropdownOption[];
}

interface ISolFilterKeys {
  ba: string;
  product: string;
  project: string;
  subject: string;
  stock: string;
  // school: string;
}

interface INewReqResponseData {
  ID: number;
  ba: string;
  product: string;
  project: string;
  subject: string;
  location: any;
  stockName: string;
  // school: string;
  availability: string;
  baValidation: boolean;
  productValidation: boolean;
  projectValidation: boolean;
  subjectValidation: boolean;
  locationValidation: boolean;
  // schoolValidation: boolean;
  stockNameValidation: boolean;
  availabilityValidation: boolean;

  overAllValidation: boolean;
  duplicateValidation: boolean;
}

interface IAddReqResponseData {
  school: any;
  requestFrom: number;
  requestCount: string;
  totalCount: number;
  stockName: string;
  stockID: number;

  requestFromValidation: boolean;
  requestCountValidation: boolean;

  overAllValidation: boolean;
}

interface IReqResponseData {
  ID: number;
  qty: any;
  status: string;
  comments: string;
}

interface IMasterProductList {
  ID: number;
  Title: string;
}

let columnSortArr: IData[] = [];
let columnSortMasterArr: IData[] = [];

const StockList = (props: IProps): JSX.Element => {
  // Variable-Declaration-Section Starts
  const sharepointWeb: IWeb = Web(props.URL);
  const stockListName: string = "Stock List";
  const requestListName: string = "Stock Request List";
  const solAllitems: IData[] = [];
  const solResReqData: IRequestData[] = [];
  const allPeoples: IPeoplelist[] = props.peopleList;
  const ListNameURL = props.WeblistURL;

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

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

  const solColumns: IColumn[] = props.isAdmin
    ? [
        {
          key: "column1",
          name: "Business area",
          fieldName: "BA",
          minWidth: 110,
          maxWidth: 110,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) =>
            BAacronymsCollection.filter((ba) => {
              return ba.Name == item.BA;
            })[0].ShortName,
        },
        {
          key: "column2",
          name: "Project",
          fieldName: "Project",
          minWidth: 150,
          maxWidth: 250,

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
                content={item.Project}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Project}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "column3",
          name: "Product",
          fieldName: "Product",
          minWidth: 150,
          maxWidth: 250,

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
                content={item.Product}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Product}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "column4",
          name: "Subject",
          fieldName: "Subject",
          minWidth: 150,
          maxWidth: 200,

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
                content={item.Subject}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Subject}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "column5",
          name: "Location",
          fieldName: "Location",
          minWidth: 150,
          maxWidth: 200,

          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => {
            let location = item.Location ? item.Location.join(",") : null;
            return (
              <>
                <TooltipHost
                  id={item.ID}
                  content={location}
                  overflowMode={TooltipOverflowMode.Parent}
                >
                  <span aria-describedby={item.ID}>{location}</span>
                </TooltipHost>
              </>
            );
          },
        },
        {
          key: "column6",
          name: "Name of stock",
          fieldName: "Stocks",
          minWidth: 200,
          maxWidth: 400,

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
                content={item.Stocks}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Stocks}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "column7",
          name: "Availability",
          fieldName: "Availability",
          minWidth: 100,
          maxWidth: 150,

          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              <div
                style={{
                  display: "flex",
                  // width: "35%",
                  justifyContent: "center",
                }}
              >
                <TooltipHost
                  id={item.ID}
                  content={item.Availability}
                  overflowMode={TooltipOverflowMode.Parent}
                >
                  <span aria-describedby={item.ID}>
                    {item.Availability > 0 ? item.Availability : "N/A"}
                  </span>
                </TooltipHost>
              </div>
              {/* {props.isAdmin ? (
            <div
              title="Edit Request"
              style={{
                display: "flex",
                // justifyContent: "end",
                // alignItems: "end",
                flexWrap: "wrap",
                width: 50,
              }}
            >
              <Icon
                iconName="Edit"
                className={solIconStyleClass.addRequestIcon}
                onClick={(): void => {
                  setSolNewRequestResponseData({
                    ID: item.ID,
                    ba: item.BA,
                    product: item.Product,
                    project: item.Project,
                    subject: item.Subject,
                    stockName: item.Stocks,
                    availability: item.Availability,
                    // school: item.School,
                    baValidation: false,
                    productValidation: false,
                    projectValidation: false,
                    subjectValidation: false,
                    // schoolValidation: false,
                    stockNameValidation: false,
                    availabilityValidation: false,

                    overAllValidation: false,
                  });
                  setSolPopup("updateRequestPopup");
                }}
              />
            </div>
          ) : null} */}
            </>
          ),
        },
        {
          key: "column8",
          name: "Approvals",
          fieldName: "Approvals",
          minWidth: 80,
          maxWidth: 100,

          onRender: (item) => (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                width: 50,
              }}
            >
              <div title={`${item.Request} pending request`}>
                <Icon
                  // iconName="CheckList"
                  iconName="OpenEnrollment"
                  className={solIconStyleClass.reqestListIcon}
                  onClick={(): void => {
                    // if (item.Request > 0) {
                    //   solGetRequestData(item.ID);
                    // }
                    solGetRequestData(item.ID, "responsesPopup");
                  }}
                />
              </div>
              <div
                title={`${item.Request} pending request`}
                style={{
                  background: "#038387",
                  color: "#fff",
                  width: 21,
                  height: 20,
                  display: "inline-flex",
                  alignItems: "center",
                  justifyContent: "center",
                  borderRadius: "50%",
                  marginLeft: 5,
                  cursor: "default",
                }}
              >
                {item.Request}
              </div>
            </div>
          ),
        },
        {
          key: "column9",
          name: "Request",
          fieldName: "Request",
          minWidth: 80,
          maxWidth: 100,

          onRender: (item) => (
            <div style={{ display: "flex" }}>
              <div
                title="Add Request"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  flexWrap: "wrap",
                  // width: props.isAdmin ? null : 50,
                  width: 50,
                }}
              >
                <Icon
                  iconName="CirclePlus"
                  className={solIconStyleClass.addRequestIcon}
                  onClick={(): void => {
                    solAddRequestResponseData.stockName = item.Stocks;
                    solAddRequestResponseData.totalCount = item.Availability;
                    solAddRequestResponseData.stockID = item.ID;

                    setSolAddRequestResponseData({
                      ...solAddRequestResponseData,
                    });
                    setSolPopup("addRequestPopup");
                  }}
                />
              </div>
            </div>
          ),
        },
        {
          key: "column10",
          name: "Action",
          fieldName: "",
          minWidth: 80,
          maxWidth: 100,

          onRender: (item) => (
            <>
              <Icon
                iconName="Edit"
                title="Edit deliverable"
                className={solIconStyleClass.edit}
                onClick={() => {
                  setSolNewRequestResponseData({
                    ID: item.ID,
                    ba: item.BA,
                    product: item.Product,
                    project: item.Project,
                    subject: item.Subject,
                    stockName: item.Stocks,
                    location: item.Location,
                    availability: item.Availability,
                    // school: item.School,
                    baValidation: false,
                    productValidation: false,
                    projectValidation: false,
                    subjectValidation: false,
                    locationValidation: false,
                    stockNameValidation: false,
                    availabilityValidation: false,

                    overAllValidation: false,
                    duplicateValidation: false,
                  });
                  setSolPopup("updateRequestPopup");
                }}
              />
              <Icon
                iconName="Delete"
                title="Delete deliverable"
                className={solIconStyleClass.delete}
                onClick={() => {
                  setSolLoader("");
                  setsolDeletePopup({
                    condition: true,
                    targetId: item.ID,
                  });
                }}
              />
            </>
          ),
        },
      ]
    : [
        {
          key: "column1",
          name: "Business area",
          fieldName: "BA",
          minWidth: 110,
          maxWidth: 110,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) =>
            BAacronymsCollection.filter((ba) => {
              return ba.Name == item.BA;
            })[0].ShortName,
        },
        {
          key: "column2",
          name: "Project",
          fieldName: "Project",
          minWidth: 150,
          maxWidth: 250,

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
                content={item.Project}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Project}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "column3",
          name: "Product",
          fieldName: "Product",
          minWidth: 150,
          maxWidth: 250,

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
                content={item.Product}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Product}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "column4",
          name: "Subject",
          fieldName: "Subject",
          minWidth: 150,
          maxWidth: 250,

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
                content={item.Subject}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Subject}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "column5",
          name: "Name of stock",
          fieldName: "Stocks",
          minWidth: 200,
          maxWidth: 400,

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
                content={item.Stocks}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Stocks}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "column6",
          name: "Availability",
          fieldName: "Availability",
          minWidth: 100,
          maxWidth: 100,

          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              <div
                style={{
                  display: "flex",
                  // width: "35%",
                  justifyContent: "center",
                }}
              >
                <TooltipHost
                  id={item.ID}
                  content={item.Availability}
                  overflowMode={TooltipOverflowMode.Parent}
                >
                  <span aria-describedby={item.ID}>
                    {item.Availability > 0 ? item.Availability : "N/A"}
                  </span>
                </TooltipHost>
              </div>
              {/* {props.isAdmin ? (
            <div
              title="Edit Request"
              style={{
                display: "flex",
                // justifyContent: "end",
                // alignItems: "end",
                flexWrap: "wrap",
                width: 50,
              }}
            >
              <Icon
                iconName="Edit"
                className={solIconStyleClass.addRequestIcon}
                onClick={(): void => {
                  setSolNewRequestResponseData({
                    ID: item.ID,
                    ba: item.BA,
                    product: item.Product,
                    project: item.Project,
                    subject: item.Subject,
                    stockName: item.Stocks,
                    availability: item.Availability,
                    // school: item.School,
                    baValidation: false,
                    productValidation: false,
                    projectValidation: false,
                    subjectValidation: false,
                    // schoolValidation: false,
                    stockNameValidation: false,
                    availabilityValidation: false,

                    overAllValidation: false,
                  });
                  setSolPopup("updateRequestPopup");
                }}
              />
            </div>
          ) : null} */}
            </>
          ),
        },
        {
          key: "column7",
          name: "Approvals",
          fieldName: "Approvals",
          minWidth: 80,
          maxWidth: 100,

          onRender: (item) => (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                width: 50,
              }}
            >
              <div title={`${item.Request} pending request`}>
                <Icon
                  // iconName="CheckList"
                  iconName="OpenEnrollment"
                  className={solIconStyleClass.reqestListIcon}
                  onClick={(): void => {
                    // if (item.Request > 0) {
                    //   solGetRequestData(item.ID);
                    // }
                    solGetRequestData(item.ID, "responsesPopup");
                  }}
                />
              </div>
              <div
                title={`${item.Request} pending request`}
                style={{
                  background: "#038387",
                  color: "#fff",
                  width: 21,
                  height: 20,
                  display: "inline-flex",
                  alignItems: "center",
                  justifyContent: "center",
                  borderRadius: "50%",
                  marginLeft: 5,
                  cursor: "default",
                }}
              >
                {item.Request}
              </div>
            </div>
          ),
        },
        {
          key: "column8",
          name: "Request",
          fieldName: "Request",
          minWidth: 80,
          maxWidth: 100,

          onRender: (item) => (
            <div style={{ display: "flex" }}>
              <div
                title="Add Request"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  flexWrap: "wrap",
                  // width: props.isAdmin ? null : 50,
                  width: 50,
                }}
              >
                <Icon
                  iconName="CirclePlus"
                  className={solIconStyleClass.addRequestIcon}
                  onClick={(): void => {
                    solAddRequestResponseData.stockName = item.Stocks;
                    solAddRequestResponseData.totalCount = item.Availability;
                    solAddRequestResponseData.stockID = item.ID;

                    setSolAddRequestResponseData({
                      ...solAddRequestResponseData,
                    });
                    setSolPopup("addRequestPopup");
                  }}
                />
              </div>
            </div>
          ),
        },
      ];
  const solRequestsColumns: IColumn[] = [
    {
      key: "column1",
      name: "From",
      fieldName: "From",
      minWidth: 260,
      maxWidth: 260,

      onRender: (item) => (
        <div style={{ display: "flex" }}>
          {item.requestFromEmail ? (
            <Persona
              size={PersonaSize.size32}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.requestFromEmail}`
              }
            />
          ) : null}
          <Label>{item.requestFromName}</Label>
        </div>
      ),
    },
    {
      key: "column2",
      name: "School",
      fieldName: "school",
      minWidth: 90,
      maxWidth: 150,
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.school}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.school}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "column3",
      name: "Created on",
      fieldName: "CreatedOn",
      minWidth: 90,
      maxWidth: 90,
      onRender: (item) => <>{moment(item.createdDate).format("DD/MM/YYYY")}</>,
    },
    {
      key: "column4",
      name: "Count",
      fieldName: "Count",
      minWidth: 70,
      maxWidth: 70,

      onRender: (item) => (
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            width: "60%",
          }}
        >
          <div>{item.requestQuantity}</div>
        </div>
        // <>
        //   <TooltipHost
        //     id={item.ID}
        //     content={item.requestQuantity}
        //     overflowMode={TooltipOverflowMode.Parent}
        //   >
        //     <span aria-describedby={item.ID}>{item.requestQuantity}</span>
        //   </TooltipHost>
        // </>
      ),
    },
    {
      key: "column5",
      name: "Actions",
      fieldName: "Actions",
      minWidth: 80,
      maxWidth: 80,

      onRender: (item) =>
        item.requestStatus != "Pending" ? (
          <div
            style={{
              backgroundColor:
                item.requestStatus == "Approved"
                  ? "#36b04b"
                  : item.requestStatus == "Rejected"
                  ? "#b80000"
                  : "",
              color: "#fff",
              padding: "5px 10px",
              borderRadius: 25,
              fontWeight: 600,
            }}
          >
            {item.requestStatus}
          </div>
        ) : (
          <>
            {solRequestResponseData.filter((res) => {
              return res.ID == item.ID;
            }).length > 0 ? (
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  width: "60%",
                }}
              >
                <div style={{ marginRight: 10 }}>
                  {
                    solRequestResponseData.filter((res) => {
                      return res.ID == item.ID;
                    })[0].status
                  }
                </div>
                <Icon
                  iconName="Refresh"
                  title="Click to reset"
                  className={solIconStyleClass.undoIcon}
                  onClick={(): void => {
                    let _tempReqResponseData: IReqResponseData[] = [
                      ...solRequestResponseData,
                    ];
                    let targetIndex: number = _tempReqResponseData.findIndex(
                      (res) => res.ID == item.ID
                    );
                    _tempReqResponseData.splice(targetIndex, 1);
                    setSolRequestResponseData([..._tempReqResponseData]);
                  }}
                />
              </div>
            ) : (
              <div style={{ display: "flex" }}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    flexWrap: "wrap",
                    // width: 50,
                  }}
                >
                  <Icon
                    iconName="CheckMark"
                    className={
                      props.isAdmin
                        ? solIconStyleClass.approveIcon
                        : solIconStyleClass.approveIconDisabled
                    }
                    onClick={(): void => {
                      if (props.isAdmin) {
                        let _tempReqResponseData: IReqResponseData[] = [
                          ...solRequestResponseData,
                        ];
                        _tempReqResponseData.push({
                          ID: item.ID,
                          qty: item.requestQuantity,
                          status: "Approved",
                          comments: "",
                        });
                        setSolRequestResponseData([..._tempReqResponseData]);
                      }
                    }}
                  />
                </div>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    flexWrap: "wrap",
                    width: 50,
                  }}
                >
                  <Icon
                    iconName="Cancel"
                    className={
                      props.isAdmin
                        ? solIconStyleClass.rejectIcon
                        : solIconStyleClass.rejectIconDisabled
                    }
                    onClick={(): void => {
                      if (props.isAdmin) {
                        let _tempReqResponseData: IReqResponseData[] = [
                          ...solRequestResponseData,
                        ];
                        _tempReqResponseData.push({
                          ID: item.ID,
                          qty: item.requestQuantity,
                          status: "Rejected",
                          comments: "",
                        });
                        setSolRequestResponseData([..._tempReqResponseData]);
                      }
                    }}
                  />
                </div>
              </div>
            )}
          </>
        ),
    },
    {
      key: "column6",
      name: "Comments",
      fieldName: "Comments",
      minWidth: 255,
      maxWidth: 300,
      onRender: (item) =>
        item.requestStatus != "Pending" ? (
          <Label
            title={
              item.comments && item.comments.length > 30 ? item.comments : null
            }
          >
            {item.comments && item.comments.length > 30
              ? item.comments.slice(0, 30) + "..."
              : item.comments}
          </Label>
        ) : (
          //
          <>
            <TextField
              styles={
                solRequestResponseData.filter((res) => {
                  return res.ID == item.ID;
                }).length > 0
                  ? solModalRequestsTxtBoxStyles
                  : solModalRequestsReadOnlyTxtBoxStyles
              }
              value={
                solRequestResponseData.filter((res) => {
                  return res.ID == item.ID;
                }).length > 0
                  ? solRequestResponseData.filter((res) => {
                      return res.ID == item.ID;
                    })[0].comments
                  : ""
              }
              readOnly={
                solRequestResponseData.filter((res) => {
                  return res.ID == item.ID;
                }).length > 0
                  ? false
                  : true
              }
              onChange={(e, value: string) => {
                let tempSolResponseData = [...solRequestResponseData];
                tempSolResponseData.filter((res) => {
                  return res.ID == item.ID;
                })[0].comments = value;
                setSolRequestResponseData([...tempSolResponseData]);
              }}
            />
          </>
        ),
    },
  ];
  const solDrpDwnOptns: ISolDrpdwn = {
    baOptns: [{ key: "All", text: "All" }],
    ProductOptns: [{ key: "All", text: "All" }],
    ProjectOptns: [{ key: "All", text: "All" }],
    SubjectOptns: [{ key: "All", text: "All" }],
    LocationOptns: [{ key: "All", text: "All" }],
    schoolOptns: [{ key: "All", text: "All" }],
  };
  const solAllDrpDwnOptns: ISolDrpdwn = {
    baOptns: [],
    ProductOptns: [],
    ProjectOptns: [],
    SubjectOptns: [],
    LocationOptns: [],
    schoolOptns: [],
  };
  const solFilterKeys: ISolFilterKeys = {
    ba: "All",
    product: "All",
    project: "All",
    subject: "All",
    stock: "",
    // school: "All",
  };
  const solNewReqResponseData: INewReqResponseData = {
    ID: null,
    ba: "",
    product: "",
    project: "",
    subject: "",
    stockName: "",
    location: [],
    availability: "",
    // school: "",
    baValidation: false,
    productValidation: false,
    projectValidation: false,
    subjectValidation: false,
    locationValidation: false,
    stockNameValidation: false,
    availabilityValidation: false,

    overAllValidation: false,
    duplicateValidation: false,
  };
  const solAddReqResponseData: IAddReqResponseData = {
    requestFrom: null,
    requestCount: "",
    totalCount: null,
    stockName: "",
    stockID: null,
    school: null,

    requestFromValidation: false,
    requestCountValidation: false,

    overAllValidation: false,
  };

  let currentpage: number = 1;
  let totalPageItems: number = 10;
  // Variable-Declaration-Section Ends
  // Styles-Section Starts
  const solDetailsListStyles: Partial<IDetailsListStyles> = {
    root: {
      overflowX: "none",
      width: "100%",
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
    contentWrapper: {
      // ".ms-DetailsRow-cell": {
      //   paddingBottom: "0 !important",
      // },
    },
  };
  const solModalDetailsListStyles: Partial<IDetailsListStyles> = {
    root: {
      overflowX: "none",
      selectors: {
        ".ms-DetailsHeader-cellTitle": {
          // justifyContent: "center !important",
        },
        ".ms-DetailsRow-cell": {
          height: "45",
          display: "flex",
          alignItems: "center",
          // justifyContent: "center",
        },
      },
    },
    headerWrapper: {},
    contentWrapper: {
      height: 135,
      marginBottom: 15,
      overflowX: "hidden",
      overflowY: "auto",
      // ".ms-DetailsRow-cell": {
      //   paddingBottom: "0 !important",
      // },
    },
  };
  const solModalDetailsListNoDataStyles: Partial<IDetailsListStyles> = {
    root: {
      overflowX: "none",
      selectors: {
        ".ms-DetailsHeader-cellTitle": {
          // justifyContent: "center !important",
        },
        ".ms-DetailsRow-cell": {
          height: "45",
          display: "flex",
          alignItems: "center",
          // justifyContent: "center",
        },
      },
    },
    headerWrapper: {},
    contentWrapper: {
      display: "none",
    },
  };
  const solLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 165,
      marginTop: 5,
      marginRight: 10,
      fontSize: 13,
      color: "#323130",
    },
  };
  const solModalBoxLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 250,
      marginTop: 5,
      fontSize: 13,
      color: "#323130",
    },
  };
  const solModalBoxReadOnlyLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 200,
      marginTop: 5,
      fontSize: 13,
      color: "#ababab",
    },
  };
  const solSearchBoxStyles: Partial<ISearchBoxStyles> = {
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
  const solActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
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
  const solDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 170,
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
    callout: {
      maxHeight: "300px",
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
  const solActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 170,
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
    callout: {
      maxHeight: "300px",
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
  const solModalReadOnlyDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 200,
      marginTop: 5,
      // marginRight: "20px",
      backgroundColor: "#fff",
      borderRadius: 3,
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      color: "#000",
      border: "1px solid #ababab",
      borderRadius: 3,
    },
    callout: {
      maxHeight: "300px",
    },
    dropdownItem: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000", display: "none" },
  };
  const solModalDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 250,
      marginTop: 5,
      // marginRight: "20px",
      backgroundColor: "#fff",
      borderRadius: 3,
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      color: "#000",
      border: "1px solid #000",
      borderRadius: 3,
    },
    callout: {
      maxHeight: "300px",
    },
    dropdownItem: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const solModalErrorDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 250,
      marginTop: 5,
      // marginRight: "20px",
      backgroundColor: "#fff",
      borderRadius: 3,
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      color: "#000",
      border: "2px solid #f00",
      borderRadius: 3,
    },
    callout: {
      maxHeight: "300px",
    },
    dropdownItem: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const solModalReadOnlyTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 200,
      marginTop: 5,
      userSelect: "none !important",
      cursor: "default !important",
    },
    fieldGroup: {
      border: "1px solid #ababab",
      userSelect: "none !important",
      cursor: "default !important",
    },
    field: {
      fontSize: 12,
      userSelect: "none !important",
      cursor: "default !important",
    },
  };
  const solModalTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 250,
      marginTop: 5,
    },
    fieldGroup: { border: "1px solid #000" },
    field: { fontSize: 12 },
  };
  const solModalErrorTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 250,
      marginTop: 5,
    },
    fieldGroup: { border: "2px solid #f00" },
    field: { fontSize: 12 },
  };
  const solModalRequestsReadOnlyTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 240,
      userSelect: "none !important",
      cursor: "default !important",
    },
    fieldGroup: {
      border: "1px solid #ababab",
      userSelect: "none !important",
      cursor: "default !important",
      backgroundColor: "rgb(231,231,231) !important",
    },
    field: {
      fontSize: 12,
      userSelect: "none !important",
      cursor: "default !important",
    },
  };
  const solModalRequestsTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: { width: 240 },
    fieldGroup: { border: "1px solid #000" },
    field: { fontSize: 12 },
  };
  const solPivotStyles: Partial<IPivotStyles> = {
    link: {
      marginLeft: "-7px",
      marginBottom: 10,
    },
  };
  const soliconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const solIconStyleClass = mergeStyleSets({
    delete: [{ color: "#CB1E06", margin: "0 0px" }, soliconStyle],
    edit: [{ color: "#2392B2", margin: "0 7px" }, soliconStyle],
    addRequestIcon: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#2392B2",
        cursor: "pointer",
      },
    ],
    reqestListIcon: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#2392B2",
        cursor: "pointer",
      },
    ],
    approveIcon: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#00a71c",
        cursor: "pointer",
      },
    ],
    approveIconDisabled: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#ababab",
        cursor: "not-allowed",
      },
    ],
    rejectIcon: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#ff0000",
        cursor: "pointer",
      },
    ],
    rejectIconDisabled: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#ababab",
        cursor: "not-allowed",
      },
    ],
    undoIcon: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#2392B2",
        cursor: "pointer",
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
        marginTop: 36.5,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
  });
  // Styles-Section Ends
  // States-Declaration Starts
  const [solReRender, setSolReRender] = useState<boolean>(true);
  const [currentUser, setCurrentUser] = useState<ICurrentuser>();
  const [masterProductList, setMasterProductList] = useState<
    IMasterProductList[]
  >([]);
  const [masterSchoolList, setMasterSchoolList] = useState<
    IMasterProductList[]
  >([]);
  const [solUnsortMasterData, setSolUnsortMasterData] =
    useState<IData[]>(solAllitems);
  const [solMasterData, setSolMasterData] = useState<IData[]>(solAllitems);
  const [solData, setSolData] = useState<IData[]>(solAllitems);
  const [solDisplayData, setSolDisplayData] = useState<IData[]>([]);
  const [solcurrentPage, setSolCurrentPage] = useState<number>(currentpage);
  const [solDropDownOptions, setSolDropDownOptions] =
    useState<ISolDrpdwn>(solDrpDwnOptns);
  const [solAllDropDownOptions, setSolAllDropDownOptions] =
    useState<ISolDrpdwn>(solDrpDwnOptns);
  const [solFilters, setSolFilters] = useState<ISolFilterKeys>(solFilterKeys);

  const [solNewRequestResponseData, setSolNewRequestResponseData] =
    useState<INewReqResponseData>(solNewReqResponseData);
  const [solAddRequestResponseData, setSolAddRequestResponseData] =
    useState<IAddReqResponseData>(solAddReqResponseData);
  const [solRequestResponseData, setSolRequestResponseData] = useState<
    IReqResponseData[]
  >([]);

  const [solRequestData, setSolRequestData] =
    useState<IRequestData[]>(solResReqData);

  const [solPopup, setSolPopup] = useState<string>("");

  const [solLoader, setSolLoader] = useState<string>("noLoader");
  const [solMasterColumns, setSolMasterColumns] =
    useState<IColumn[]>(solColumns);
  const [solDeletePopup, setsolDeletePopup] = useState({
    condition: false,
    targetId: 0,
  });

  // States-Declaration Ends
  //Function-Section Starts
  const solGetCurrentUserDetails = (): void => {
    sharepointWeb.currentUser
      .get()
      .then((user) => {
        let solCurrentUser: ICurrentuser = {
          Name: user.Title,
          Email: user.Email,
          Id: user.Id,
        };
        setCurrentUser(solCurrentUser);
      })
      .catch((err) => {
        solErrorFunction(err, "solGetCurrentUserDetails");
      });
  };
  const solGetData = (): void => {
    sharepointWeb.lists
      .getByTitle(stockListName)
      .items.select(
        "*",
        "Product/ProductVersion",
        "Product/Title",
        "Product/Id"
        // "School/Title",
        // "School/Id"
      )
      .expand("Product")
      .orderBy("Modified", false)
      .top(5000)
      .get()
      .then((items) => {
        console.log(items);
        items.forEach((item) => {
          solAllitems.push({
            ID: item.Id ? item.Id : "",
            BA: item.BA ? item.BA : "",
            Product: item.ProductId
              ? item.Product.Title + " " + item.Product.ProductVersion
              : "",
            Project: item.Project ? item.Project : "",
            Subject: item.Subject ? item.Subject : "",
            Stocks: item.Stocks ? item.Stocks : "",
            Location: item.Location ? item.Location : [],
            // School: item.SchoolId ? item.School.Title : "",
            Availability: item.Title ? item.Title : null,
            Request: item.Request ? item.Request : null,
          });
        });
        console.log(solAllitems);

        solGetAllDrpDwnOptions(solAllitems);
        paginateFunction(1, solAllitems);

        setSolUnsortMasterData([...solAllitems]);
        columnSortArr = solAllitems;
        setSolData([...solAllitems]);
        columnSortMasterArr = solAllitems;
        setSolMasterData([...solAllitems]);
        setSolLoader("noLoader");
      })
      .catch((err) => {
        solErrorFunction(err, "solGetData");
      });
  };
  const solGetRequestData = (targetID: number, popupName: string): void => {
    console.log(popupName);
    sharepointWeb.lists
      .getByTitle(requestListName)
      .items.select(
        "*",
        "RequestFrom/Title",
        "RequestFrom/Id",
        "RequestFrom/EMail",
        "School/Title",
        "School/Id"
      )
      .expand("RequestFrom", "School")
      .orderBy("Modified", false)
      .get()
      .then((items) => {
        let _reqResponse: IRequestData[] = [];
        items = items.filter((item) => {
          return parseInt(item.StockID) == targetID;
        });
        console.log(items);
        if (items.length > 0) {
          items.forEach((item) => {
            _reqResponse.push({
              ID: item.Id,
              requestFromName: item.RequestFrom.Title,
              requestFromId: item.RequestFrom.Id,
              requestFromEmail: item.RequestFrom.EMail,
              requestQuantity: item.RequestQuantity,
              requestStatus: item.RequestStatus,
              // stockName: item.StockName,
              stockID: item.StockID,
              createdDate: item.Created,
              comments: item.Comments,
              school: item.SchoolId ? item.School.Title : "",
            });
          });
        } else {
          _reqResponse.push({
            ID: null,
            requestFromName: "",
            requestFromId: null,
            requestFromEmail: "",
            requestQuantity: null,
            requestStatus: "",
            // stockName: item.StockName,
            stockID: targetID,
            createdDate: null,
            comments: "",
            school: null,
          });
          // setSolRequestData([..._reqResponse]);
          // setSolPopup(popupName);
          // console.log(_reqResponse);
        }
        setSolRequestData([..._reqResponse]);
        setSolPopup(popupName);
        console.log(_reqResponse);
      })
      .catch((err) => {
        solErrorFunction(err, "solGetRequestData");
      });
  };
  const solGetAllOptions = (): void => {
    sharepointWeb.lists
      .getByTitle(stockListName)
      .fields.getByInternalNameOrTitle("BA")
      .get()
      .then((response: any) => {
        response.Choices.forEach((choice: string) => {
          if (choice) {
            solAllDrpDwnOptns.baOptns.push({
              key: choice,
              text: choice,
            });
          }
        });
        sharepointWeb.lists
          .getByTitle("Master Product List")
          .items.filter("IsDeleted ne 1")
          .top(5000)
          .get()
          .then((_response: any) => {
            _response.forEach((choice: string) => {
              if (choice) {
                solAllDrpDwnOptns.ProductOptns.push({
                  key: choice["Title"] + " " + choice["ProductVersion"],
                  text: choice["Title"] + " " + choice["ProductVersion"],
                });
                masterProductList.push({
                  ID: choice["ID"],
                  Title: choice["Title"] + " " + choice["ProductVersion"],
                });
              }
            });
            sharepointWeb.lists
              .getByTitle("Stock Config List")
              .items.get()
              .then((_response: any) => {
                _response.forEach((choice: string) => {
                  if (choice) {
                    solAllDrpDwnOptns.schoolOptns.push({
                      key: choice["Title"],
                      text: choice["Title"],
                    });
                    masterSchoolList.push({
                      ID: choice["ID"],
                      Title: choice["Title"],
                    });
                  }
                });

                sharepointWeb.lists
                  .getByTitle(stockListName)
                  .fields.getByInternalNameOrTitle("Subject")
                  .get()
                  .then((response: any) => {
                    response.Choices.forEach((choice: string) => {
                      if (choice) {
                        solAllDrpDwnOptns.SubjectOptns.push({
                          key: choice,
                          text: choice,
                        });
                      }
                    });
                    sharepointWeb.lists
                      .getByTitle(ListNameURL)
                      .items.top(5000)
                      .orderBy("Modified", false)
                      .get()
                      .then((Items) => {
                        Items.forEach((arr) => {
                          if (
                            solAllDrpDwnOptns.ProjectOptns.findIndex((prj) => {
                              return prj.key == arr.Title;
                            }) == -1 &&
                            arr.Title
                          ) {
                            solAllDrpDwnOptns.ProjectOptns.push({
                              key: arr.Title + " " + arr.ProjectVersion,
                              text: arr.Title + " " + arr.ProjectVersion,
                            });
                          }
                        });
                        sharepointWeb.lists
                          .getByTitle(stockListName)
                          .fields.getByInternalNameOrTitle("Location")
                          .get()
                          .then((response: any) => {
                            response.Choices.forEach((choice: string) => {
                              if (choice) {
                                solAllDrpDwnOptns.LocationOptns.push({
                                  key: choice,
                                  text: choice,
                                });
                              }
                            });
                            let unsortedFilterKeys: ISolDrpdwn =
                              solSortingAllDrpDwnOptns(solAllDrpDwnOptns);
                            setSolAllDropDownOptions(unsortedFilterKeys);
                            setMasterProductList([...masterProductList]);
                            setMasterSchoolList([...masterSchoolList]);
                          })
                          .catch((err) => {
                            solErrorFunction(
                              err,
                              "solGetAllOptions-getStockList"
                            );
                          });
                        // .then(() => {

                        // })
                        // .catch((err)=>{solErrorFunction(err)});
                      })
                      .catch((err) => {
                        solErrorFunction(err, "solGetAllOptions-APItems");
                      });
                  })
                  .catch((err) => {
                    solErrorFunction(err, "solGetAllOptions-Subject");
                  });
              })
              .catch((err) => {
                solErrorFunction(err, "solGetAllOptions-getStockConfigList");
              });
          })
          .catch((err) => {
            solErrorFunction(err, "solGetAllOptions-getMasterProductList");
          });
      })
      .catch((err) => {
        solErrorFunction(err, "solGetAllOptions-BA");
      });
  };
  const solSortingAllDrpDwnOptns = (
    unsortedFilterKeys: ISolDrpdwn
  ): ISolDrpdwn => {
    const sortFilterKeys = (a, b): number => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    unsortedFilterKeys.baOptns.sort(sortFilterKeys);
    unsortedFilterKeys.ProductOptns.sort(sortFilterKeys);
    unsortedFilterKeys.schoolOptns.sort(sortFilterKeys);
    unsortedFilterKeys.SubjectOptns.sort(sortFilterKeys);
    unsortedFilterKeys.ProjectOptns.sort(sortFilterKeys);

    return unsortedFilterKeys;
  };
  const solGetAllDrpDwnOptions = (allItems: IData[]): void => {
    allItems.forEach((item: any) => {
      if (
        solDrpDwnOptns.baOptns.findIndex((baOptn) => {
          return baOptn.key == item.BA;
        }) == -1 &&
        item.BA
      ) {
        solDrpDwnOptns.baOptns.push({
          key: item.BA,
          text: item.BA,
        });
      }

      if (
        solDrpDwnOptns.ProductOptns.findIndex((ProductOptn) => {
          return ProductOptn.key == item.Product;
        }) == -1 &&
        item.Product
      ) {
        solDrpDwnOptns.ProductOptns.push({
          key: item.Product,
          text: item.Product,
        });
      }
      if (
        solDrpDwnOptns.schoolOptns.findIndex((SchoolOptn) => {
          return SchoolOptn.key == item.School;
        }) == -1 &&
        item.School
      ) {
        solDrpDwnOptns.schoolOptns.push({
          key: item.School,
          text: item.School,
        });
      }
    });

    let unsortedFilterKeys: ISolDrpdwn = solSortingDrpDwnOptns(solDrpDwnOptns);
    setSolDropDownOptions(unsortedFilterKeys);
  };
  const solSortingDrpDwnOptns = (
    unsortedFilterKeys: ISolDrpdwn
  ): ISolDrpdwn => {
    const sortFilterKeys = (a, b): number => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    unsortedFilterKeys.baOptns.shift();
    unsortedFilterKeys.baOptns.sort(sortFilterKeys);
    unsortedFilterKeys.baOptns.unshift({ key: "All", text: "All" });

    unsortedFilterKeys.ProductOptns.shift();
    unsortedFilterKeys.ProductOptns.sort(sortFilterKeys);
    unsortedFilterKeys.ProductOptns.unshift({ key: "All", text: "All" });

    unsortedFilterKeys.schoolOptns.shift();
    unsortedFilterKeys.schoolOptns.sort(sortFilterKeys);
    unsortedFilterKeys.schoolOptns.unshift({ key: "All", text: "All" });

    return unsortedFilterKeys;
  };
  const solListFilter = (key: string, option: any): void => {
    let arrBeforeFilter: IData[] = [...solMasterData];
    let tempFilterKeys: ISolFilterKeys = { ...solFilters };
    tempFilterKeys[key] = option;

    if (tempFilterKeys.ba != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.BA == tempFilterKeys.ba;
      });
    }

    if (tempFilterKeys.product != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Product == tempFilterKeys.product;
      });
    }
    // if (tempFilterKeys.school != "All") {
    //   arrBeforeFilter = arrBeforeFilter.filter((arr) => {
    //     return arr.School == tempFilterKeys.school;
    //   });
    // }
    if (tempFilterKeys.stock) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Stocks.toLowerCase().includes(
          tempFilterKeys.stock.toLowerCase()
        );
      });
    }

    paginateFunction(1, arrBeforeFilter);

    columnSortArr = arrBeforeFilter;
    setSolData([...columnSortArr]);
    setSolFilters({ ...tempFilterKeys });
  };
  const solListFilterByData = (allItems: IData[]): IData[] => {
    let arrBeforeFilter: IData[] = [...allItems];
    let tempFilterKeys: ISolFilterKeys = { ...solFilters };

    if (tempFilterKeys.ba != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.BA == tempFilterKeys.ba;
      });
    }

    if (tempFilterKeys.product != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Product == tempFilterKeys.product;
      });
    }
    // if (tempFilterKeys.school != "All") {
    //   arrBeforeFilter = arrBeforeFilter.filter((arr) => {
    //     return arr.School == tempFilterKeys.school;
    //   });
    // }
    if (tempFilterKeys.stock) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Stocks.toLowerCase().includes(
          tempFilterKeys.stock.toLowerCase()
        );
      });
    }

    return arrBeforeFilter;
  };
  const solDataHandler = (
    responseType: string,
    key: string,
    value: any
  ): void => {
    if (responseType == "newRequest") {
      if (key == "location") {
        if (value) {
          solNewRequestResponseData.location = value.selected
            ? [...solNewRequestResponseData.location, value.key as string]
            : solNewRequestResponseData.location.filter(
                (key) => key !== value.key
              );
          solNewRequestResponseData.location.sort();
        }
      } else {
        solNewRequestResponseData[key] = value;
      }

      // solNewRequestResponseData[key] = value;
      solNewRequestResponseData.overAllValidation = false;
      key == "ba"
        ? (solNewRequestResponseData.baValidation = false)
        : key == "product"
        ? (solNewRequestResponseData.productValidation = false)
        : key == "project"
        ? (solNewRequestResponseData.projectValidation = false)
        : key == "subject"
        ? (solNewRequestResponseData.subjectValidation = false)
        : key == "location"
        ? (solNewRequestResponseData.locationValidation = false)
        : key == "stockName"
        ? (solNewRequestResponseData.stockNameValidation = false)
        : key == "availability"
        ? (solNewRequestResponseData.availabilityValidation = false)
        : "",
        console.log(solNewRequestResponseData);
      setSolNewRequestResponseData({ ...solNewRequestResponseData });
    } else if (responseType == "addRequest") {
      solAddRequestResponseData[key] = value;
      solAddRequestResponseData.overAllValidation = false;
      key == "requestFrom"
        ? (solAddRequestResponseData.requestFromValidation = false)
        : key == "requestCount"
        ? (solAddRequestResponseData.requestCountValidation = false)
        : "",
        console.log(solAddRequestResponseData);
      setSolAddRequestResponseData({ ...solAddRequestResponseData });
    } else if (responseType == "requestResponse") {
    }
  };
  const validationFunction = (moduleType: string): void => {
    const validationCheck = (
      data: any,
      fieldName: string,
      validationName: string,
      overAllValidationState: string
    ): void => {
      if (fieldName == "availability") {
        if (!data[fieldName] || data[fieldName] == "0") {
          data[validationName] = true;

          !data[overAllValidationState]
            ? (data[overAllValidationState] = true)
            : "";
        }
      } else {
        if (!data[fieldName]) {
          data[validationName] = true;

          !data[overAllValidationState]
            ? (data[overAllValidationState] = true)
            : "";
        }
      }
    };

    let duplicateArr = solMasterData.filter((arr) => {
      if (solNewRequestResponseData.ID) {
        return (
          arr.BA == solNewRequestResponseData.ba &&
          arr.Product == solNewRequestResponseData.product &&
          arr.Project == solNewRequestResponseData.project &&
          arr.Subject == solNewRequestResponseData.subject &&
          arr.Stocks == solNewRequestResponseData.stockName &&
          arr.ID != solNewRequestResponseData.ID
        );
      } else {
        return (
          arr.BA == solNewRequestResponseData.ba &&
          arr.Product == solNewRequestResponseData.product &&
          arr.Project == solNewRequestResponseData.project &&
          arr.Subject == solNewRequestResponseData.subject &&
          arr.Stocks == solNewRequestResponseData.stockName
        );
      }
    });

    solNewRequestResponseData.duplicateValidation =
      duplicateArr.length > 0 ? true : false;

    if (moduleType == "newRequest") {
      validationCheck(
        solNewRequestResponseData,
        "ba",
        "baValidation",
        "overAllValidation"
      );
      validationCheck(
        solNewRequestResponseData,
        "product",
        "productValidation",
        "overAllValidation"
      );
      validationCheck(
        solNewRequestResponseData,
        "project",
        "projectValidation",
        "overAllValidation"
      );
      validationCheck(
        solNewRequestResponseData,
        "subject",
        "subjectValidation",
        "overAllValidation"
      );
      // validationCheck(
      //   solNewRequestResponseData,
      //   "school",
      //   "schoolValidation",
      //   "overAllValidation"
      // );
      validationCheck(
        solNewRequestResponseData,
        "stockName",
        "stockNameValidation",
        "overAllValidation"
      );
      validationCheck(
        solNewRequestResponseData,
        "availability",
        "availabilityValidation",
        "overAllValidation"
      );
      if (solNewRequestResponseData.overAllValidation) {
        console.log(solNewRequestResponseData);
        setSolNewRequestResponseData({ ...solNewRequestResponseData });
        setSolLoader("");
      } else if (duplicateArr.length > 0) {
        console.log(solNewRequestResponseData);
        setSolNewRequestResponseData({ ...solNewRequestResponseData });
        setSolLoader("");
      } else {
        console.log(solNewRequestResponseData);
        solPopup == "newRequestPopup"
          ? addNewRequestFunction(solNewRequestResponseData)
          : solPopup == "updateRequestPopup"
          ? updateAddRequestFunction(
              solNewRequestResponseData.ID,
              solMasterData
            )
          : null;
      }
    } else if (moduleType == "addRequest") {
      validationCheck(
        solAddRequestResponseData,
        "requestFrom",
        "requestFromValidation",
        "overAllValidation"
      );
      validationCheck(
        solAddRequestResponseData,
        "requestCount",
        "requestCountValidation",
        "overAllValidation"
      );
      if (solAddRequestResponseData.overAllValidation) {
        console.log(solAddRequestResponseData);
        setSolAddRequestResponseData({ ...solAddRequestResponseData });
        setSolLoader("");
      } else {
        console.log(solAddRequestResponseData);
        add_AddRequestFunction(solAddRequestResponseData);
      }
    }
  };

  const addNewRequestFunction = (_data: INewReqResponseData): void => {
    let tempMasterData: IData[] = [...solMasterData];
    let responseData: {
      BA: string;
      ProductId: number;
      Project: string;
      Subject: string;
      Stocks: string;
      Location: any;
      // SchoolId: number;
      Title: string;
      Request: string;
    } = {
      BA: _data.ba ? _data.ba : "",
      ProductId: _data.product
        ? masterProductList.filter((prod) => {
            return prod.Title == _data.product;
          })[0].ID
        : null,
      Location:
        _data.location.length > 0
          ? { results: [..._data.location] }
          : { results: [] },
      Project: _data.project ? _data.project : "",
      Subject: _data.subject ? _data.subject : "",
      Stocks: _data.stockName ? _data.stockName : "",
      // SchoolId: _data.school
      //   ? masterSchoolList.filter((prod) => {
      //       return prod.Title == _data.school;
      //     })[0].ID
      //   : null,
      Title: _data.availability ? _data.availability : "",
      Request: "0",
    };

    console.log(responseData);

    sharepointWeb.lists
      .getByTitle(stockListName)
      .items.add(responseData)
      .then((e) => {
        console.log(e.data.ID);

        tempMasterData.unshift({
          ID: e.data.ID,
          BA: _data.ba ? _data.ba : "",
          Product: _data.product ? _data.product : "",
          Project: _data.project ? _data.project : "",
          Subject: _data.subject ? _data.subject : "",
          Location: _data.location ? _data.location : [],
          // School: _data.school ? _data.school : "",
          Stocks: _data.stockName ? _data.stockName : "",
          Availability: _data.availability ? _data.availability : "",
          Request: "0",
        });

        solGetAllDrpDwnOptions(tempMasterData);

        solListFilterByData(
          solListFilterByData(tempMasterData).length > 0
            ? solListFilterByData(tempMasterData)
            : tempMasterData
        );

        paginateFunction(
          1,
          solListFilterByData(tempMasterData).length > 0
            ? solListFilterByData(tempMasterData)
            : tempMasterData
        );

        setSolUnsortMasterData([...tempMasterData]);
        columnSortArr = tempMasterData;
        setSolData([...tempMasterData]);
        columnSortMasterArr = tempMasterData;
        setSolMasterData([...tempMasterData]);
        setSolNewRequestResponseData({ ...solNewReqResponseData });
        setSolPopup("");
        setSolLoader("noLoader");
      })
      .catch((err) => {
        solErrorFunction(err, "addNewRequestFunction");
      });
  };

  const add_AddRequestFunction = (_data: IAddReqResponseData): void => {
    let responseData: {
      RequestFromId: number;
      RequestQuantity: string;
      RequestStatus: string;
      StockID: string;
      SchoolId: number;
    } = {
      RequestFromId: _data.requestFrom ? _data.requestFrom : null,
      RequestQuantity: _data.requestCount ? _data.requestCount : "",
      RequestStatus: "Pending",
      StockID: _data.stockID.toString(),
      SchoolId: _data.school
        ? masterSchoolList.filter((prod) => {
            return prod.Title == _data.school;
          })[0].ID
        : null,
    };

    console.log(responseData);

    sharepointWeb.lists
      .getByTitle(requestListName)
      .items.add(responseData)
      .then((e) => {
        console.log(e);
        updateAddRequestFunction(_data.stockID, solMasterData);
      })
      .catch((err) => {
        solErrorFunction(err, "add_AddRequestFunction");
      });
  };
  const updateAddRequestFunction = (targetId: number, _data: IData[]): void => {
    let responseData = {};
    let updatedRequestCount: string;
    let updatedAvailability: string;
    if (solPopup == "addRequestPopup") {
      updatedRequestCount = (
        parseInt(
          _data.filter((data: IData) => {
            return data.ID == targetId;
          })[0].Request
        ) + 1
      ).toString();
      responseData = {
        Request: parseInt(updatedRequestCount) > 0 ? updatedRequestCount : 0,
      };
    } else if (solPopup == "updateRequestPopup") {
      updatedAvailability = solNewRequestResponseData.availability.toString();

      responseData = {
        BA: solNewRequestResponseData.ba ? solNewRequestResponseData.ba : "",
        ProductId: solNewRequestResponseData.product
          ? masterProductList.filter((prod) => {
              return prod.Title == solNewRequestResponseData.product;
            })[0].ID
          : null,
        Location:
          solNewRequestResponseData.location.length > 0
            ? { results: [...solNewRequestResponseData.location] }
            : { results: [] },
        Project: solNewRequestResponseData.project
          ? solNewRequestResponseData.project
          : "",
        Subject: solNewRequestResponseData.subject
          ? solNewRequestResponseData.subject
          : "",
        Stocks: solNewRequestResponseData.stockName
          ? solNewRequestResponseData.stockName
          : "",
        Title: updatedAvailability,
      };
    }
    // updatedRequestCount = (parseInt(updatedRequestCount) + 1).toString();
    console.log(responseData);

    sharepointWeb.lists
      .getByTitle(stockListName)
      .items.getById(targetId)
      .update(responseData)
      .then(() => {
        let targetIndex: number = _data.findIndex((arr) => arr.ID == targetId);

        solPopup == "addRequestPopup"
          ? (_data[targetIndex].Request = updatedRequestCount)
          : solPopup == "updateRequestPopup"
          ? ((_data[targetIndex].Availability = updatedAvailability),
            (_data[targetIndex].BA = solNewRequestResponseData.ba
              ? solNewRequestResponseData.ba
              : ""),
            (_data[targetIndex].Product = solNewRequestResponseData.product
              ? solNewRequestResponseData.product
              : ""),
            (_data[targetIndex].Project = solNewRequestResponseData.project
              ? solNewRequestResponseData.project
              : ""),
            (_data[targetIndex].Subject = solNewRequestResponseData.subject
              ? solNewRequestResponseData.subject
              : ""),
            (_data[targetIndex].Stocks = solNewRequestResponseData.stockName
              ? solNewRequestResponseData.stockName
              : ""),
            (_data[targetIndex].Location = solNewRequestResponseData.location
              ? solNewRequestResponseData.location
              : []))
          : null;

        solGetAllDrpDwnOptions(_data);

        solListFilterByData(
          solListFilterByData(_data).length > 0
            ? solListFilterByData(_data)
            : _data
        );

        paginateFunction(
          1,
          solListFilterByData(_data).length > 0
            ? solListFilterByData(_data)
            : _data
        );

        setSolUnsortMasterData([..._data]);
        columnSortArr = _data;
        setSolData([..._data]);
        columnSortMasterArr = _data;
        setSolMasterData([..._data]);
        // setSolLoader("startUpLoader");
        setSolNewRequestResponseData({ ...solNewReqResponseData });
        setSolAddRequestResponseData({ ...solAddReqResponseData });
        setSolPopup("");
        setSolLoader("noLoader");
      })
      .catch((err) => {
        solErrorFunction(err, "updateAddRequestFunction");
      });
  };

  const submitRequestFunction = (_data: IReqResponseData[]): void => {
    let count: number = 0;
    _data.forEach((data: IReqResponseData) => {
      let responseData: { RequestStatus: string; Comments: string } = {
        RequestStatus: data.status,
        Comments: data.comments,
      };
      sharepointWeb.lists
        .getByTitle(requestListName)
        .items.getById(data.ID)
        .update(responseData)
        .then(() => {
          count++;
          if (count == _data.length) {
            let newAvailability: number =
              solMasterData.filter((item) => {
                return item.ID == solRequestData[0].stockID;
              })[0].Availability - calculateProposedQty(solRequestResponseData);
            let newRequestCount: number =
              solMasterData.filter((item) => {
                return item.ID == solRequestData[0].stockID;
              })[0].Request - solRequestResponseData.length;
            let _responseData: { Title: string; Request: string } = {
              Title: newAvailability.toString(),
              Request: newRequestCount < 0 ? "0" : newRequestCount.toString(),
            };
            console.log(_responseData);

            sharepointWeb.lists
              .getByTitle(stockListName)
              .items.getById(solRequestData[0].stockID)
              .update(_responseData)
              .then(() => {
                let targetData: IData = solMasterData.filter((data) => {
                  return data.ID == solRequestData[0].stockID;
                })[0];
                let targetIndex: number = solMasterData.findIndex(
                  (data) => data.ID == solRequestData[0].stockID
                );
                solMasterData.splice(targetIndex, 1);
                targetData.Availability = newAvailability;
                targetData.Request = newRequestCount;
                solMasterData.unshift(targetData);

                solGetAllDrpDwnOptions(solMasterData);
                paginateFunction(1, solMasterData);

                setSolUnsortMasterData([...solMasterData]);
                columnSortArr = solMasterData;
                setSolData([...solMasterData]);
                columnSortMasterArr = solMasterData;
                setSolMasterData([...solMasterData]);
                setSolRequestResponseData([]);
                setSolPopup("");
                setSolLoader("noLoader");
              })
              .catch((err) => {
                solErrorFunction(err, "submitRequestFunction-updateStockList");
              });
          }
        })
        .catch((err) => {
          solErrorFunction(err, "submitRequestFunction-updateRequestList");
        });
    });
  };

  const paginateFunction = (pagenumber: number, data: IData[]): void => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems: IData[] = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setSolDisplayData(paginatedItems);
      setSolCurrentPage(pagenumber);
    } else {
      setSolDisplayData([]);
      setSolCurrentPage(1);
    }
  };
  const calculateProposedQty = (
    _reqResponseData: IReqResponseData[]
  ): number => {
    let proposedQty: number = null;
    _reqResponseData.forEach((data: IReqResponseData) => {
      if (data.status == "Approved") {
        proposedQty += parseInt(data.qty);
      }
    });
    return proposedQty;
  };
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempSolColumns = solColumns;
    const newColumns: IColumn[] = tempSolColumns.slice();
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

    const newDisplayData = _copyAndSort(
      columnSortArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newMasterData = _copyAndSort(
      columnSortMasterArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setSolData([...newDisplayData]);
    setSolMasterData([...newMasterData]);
    paginateFunction(1, newDisplayData);
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
  const doesTextStartWith = (text: string, filterText: string): boolean => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };
  const DeleteSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Stock is successfully deleted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const solDeleteItem = (id: number) => {
    sharepointWeb.lists
      .getByTitle(stockListName)
      .items.getById(id)
      .delete()
      .then(() => {
        let tempUnsortMasterArr = [...solUnsortMasterData];
        let targetUnsortIndex = tempUnsortMasterArr.findIndex(
          (arr) => arr.ID == id
        );
        tempUnsortMasterArr.splice(targetUnsortIndex, 1);

        let tempMasterArr = [...solMasterData];
        let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
        tempMasterArr.splice(targetIndex, 1);

        let temp_sol_arr = [...solData];
        let targetIndexapdata = temp_sol_arr.findIndex((arr) => arr.ID == id);
        temp_sol_arr.splice(targetIndexapdata, 1);

        let temp_dis_arr = [...solData];
        let targetIndexdisdata = temp_dis_arr.findIndex((arr) => arr.ID == id);
        temp_dis_arr.splice(targetIndexdisdata, 1);

        solGetAllDrpDwnOptions(solAllitems);
        paginateFunction(1, solAllitems);

        setSolUnsortMasterData([...tempUnsortMasterArr]);
        columnSortMasterArr = tempUnsortMasterArr;
        setSolData([...temp_sol_arr]);
        columnSortArr = temp_sol_arr;
        setSolMasterData([...tempMasterArr]);
        setSolDisplayData([...temp_dis_arr]);
        setsolDeletePopup({ condition: false, targetId: 0 });
        DeleteSuccessPopup();
      })
      .catch((err) => {
        solErrorFunction(err, "solDeleteItem");
      });
  };
  const solErrorFunction = (error: string, functionName: string): void => {
    console.log(error);

    let response = {
      ComponentName: "Stock list",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setSolPopup("");
        setSolLoader("noLoader");
        ErrorPopup();
      }
    );
  };
  //Function-Section Ends
  useEffect(() => {
    setSolLoader("startUpLoader");
    solGetCurrentUserDetails();
    solGetAllOptions();
    solGetData();
  }, [solReRender]);

  return (
    <>
      <div style={{ padding: "5px 10px" }}>
        {solLoader == "startUpLoader" ? <CustomLoader /> : null}
        {/* Header-Section Starts */}
        <div
          style={{
            position: "sticky",
            top: 0,
            backgroundColor: "#fff",
            zIndex: 1,
            marginBottom: 5,
          }}
        >
          <div
            className={styles.solHeaderSection}
            style={{ paddingBottom: "5px" }}
          >
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <div className={styles.solHeader}>Stock list</div>
            </div>
            {props.isAdmin ? (
              <div style={{ display: "flex", justifyContent: "flex-start" }}>
                <div>
                  <button
                    className={styles.solAddBtn}
                    onClick={(): void => {
                      setSolPopup("newRequestPopup");
                    }}
                  >
                    Add stock
                  </button>
                </div>
              </div>
            ) : null}

            {/* Filter-Section Starts */}
            <div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  flexWrap: "wrap",
                }}
              >
                <div className={styles.ddSection}>
                  <div>
                    <Label styles={solLabelStyles}>Business area</Label>
                    <Dropdown
                      placeholder="Select an option"
                      styles={
                        solFilters.ba != "All"
                          ? solActiveDropdownStyles
                          : solDropdownStyles
                      }
                      options={solDropDownOptions.baOptns}
                      dropdownWidth={"auto"}
                      onChange={(e, option: any): void => {
                        solListFilter("ba", option["key"]);
                      }}
                      selectedKey={solFilters.ba}
                    />
                  </div>
                  <div>
                    <Label styles={solLabelStyles}>Product</Label>
                    <Dropdown
                      placeholder="Select an option"
                      styles={
                        solFilters.product != "All"
                          ? solActiveDropdownStyles
                          : solDropdownStyles
                      }
                      options={solDropDownOptions.ProductOptns}
                      dropdownWidth={"auto"}
                      onChange={(e, option: any) => {
                        solListFilter("product", option["key"]);
                      }}
                      selectedKey={solFilters.product}
                    />
                  </div>
                  {/* <div>
                    <Label styles={solLabelStyles}>School</Label>
                    <Dropdown
                      placeholder="Select an option"
                      styles={
                        solFilters.school != 'All'
                          ? solActiveDropdownStyles
                          : solDropdownStyles
                      }
                      options={solDropDownOptions.schoolOptns}
                      dropdownWidth={'auto'}
                      onChange={(e, option: any) => {
                        solListFilter('school', option['key'])
                      }}
                      selectedKey={solFilters.school}
                    />
                  </div> */}
                  <div>
                    <Label styles={solLabelStyles}>Name of stock</Label>
                    <SearchBox
                      styles={
                        solFilters.stock
                          ? solActiveSearchBoxStyles
                          : solSearchBoxStyles
                      }
                      value={solFilters.stock}
                      onChange={(e, value): void => {
                        solListFilter("stock", value);
                      }}
                    />
                  </div>
                  <div>
                    <Icon
                      iconName="Refresh"
                      title="Click to reset"
                      className={solIconStyleClass.refresh}
                      onClick={(): void => {
                        paginateFunction(1, solUnsortMasterData);
                        columnSortArr = solMasterData;
                        setSolData(solMasterData);
                        columnSortMasterArr = solMasterData;
                        setSolMasterData(solMasterData);
                        setSolMasterColumns(solColumns);
                        // solGetAllDrpDwnOptions(solMasterData);
                        setSolFilters({ ...solFilterKeys });
                      }}
                    />
                  </div>
                </div>
                <div>
                  <Label
                    style={{
                      marginTop: "25px",
                      marginLeft: "10px",
                      fontWeight: "500",
                      color: "#323130",
                      fontSize: "13px",
                    }}
                  >
                    Number of records :{" "}
                    <span style={{ color: "#038387" }}>{solData.length}</span>
                  </Label>
                </div>
              </div>
            </div>
            {/* Filter-Section Ends */}
          </div>
        </div>
        {/* Header-Section Ends */}
        {/* Body-Section Starts */}
        <div>
          {/* DetailList-Section Starts */}
          <div>
            <DetailsList
              items={solDisplayData}
              columns={solMasterColumns}
              // styles={solDetailsListStyles}
              styles={{
                root: {
                  width: "100%",
                  selectors: {
                    ".ms-DetailsRow-cell": {
                      height: 45,
                      display: "flex",
                      alignItems: "center",
                    },
                  },
                },
              }}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
            />
          </div>
          {/* DetailList-Section Ends */}
          {solData.length > 0 ? (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                margin: "10px 0",
              }}
            >
              <Pagination
                currentPage={solcurrentPage}
                totalPages={
                  solData.length > 0
                    ? Math.ceil(solData.length / totalPageItems)
                    : 1
                }
                onChange={(page): void => {
                  paginateFunction(page, solData);
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
        {/* Body-Section Ends */}
        {/* Modal-Section Starts */}
        <div>
          {/* New & Update Request-Popup Starts */}
          <div>
            {solPopup == "newRequestPopup" ||
            solPopup == "updateRequestPopup" ? (
              <Modal
                isOpen={
                  solPopup == "newRequestPopup" ||
                  solPopup == "updateRequestPopup"
                    ? true
                    : false
                }
                styles={{
                  main: {
                    // width: "50%",
                    width: 900,
                  },
                  // root: {
                  //   selectors: {
                  //     ".ms-Dialog-main": {
                  //       width: "754px",
                  //     },
                  //   },
                  // },
                }}
                isBlocking={true}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    width: "100%",
                    padding: "20px 0",
                    paddingBottom: "10px",
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
                    <Label className={styles.newRequestTitle}>
                      {solPopup == "updateRequestPopup"
                        ? "Update stock"
                        : "New stock"}
                    </Label>
                  </div>
                </div>
                <div style={{ width: "90%", margin: "auto" }}>
                  <div
                    style={{ display: "flex", justifyContent: "space-between" }}
                  >
                    <div>
                      <Label
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalBoxReadOnlyLabelStyles
                          //   :
                          solModalBoxLabelStyles
                        }
                      >
                        Business area
                      </Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalReadOnlyDropdownStyles
                          //   :
                          solNewRequestResponseData.baValidation
                            ? solModalErrorDropdownStyles
                            : solModalDropdownStyles
                        }
                        // disabled={
                        //   solPopup == "updateRequestPopup" ? true : false
                        // }
                        options={solAllDropDownOptions.baOptns}
                        dropdownWidth={"auto"}
                        onChange={(e, option: any): void => {
                          solDataHandler("newRequest", "ba", option["key"]);
                        }}
                        selectedKey={solNewRequestResponseData.ba}
                      />
                    </div>
                    <div>
                      <Label
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalBoxReadOnlyLabelStyles
                          //   :
                          solModalBoxLabelStyles
                        }
                      >
                        Project
                      </Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalReadOnlyDropdownStyles
                          //   :
                          solNewRequestResponseData.projectValidation
                            ? solModalErrorDropdownStyles
                            : solModalDropdownStyles
                        }
                        // disabled={
                        //   solPopup == "updateRequestPopup" ? true : false
                        // }
                        options={solAllDropDownOptions.ProjectOptns}
                        dropdownWidth={"auto"}
                        onChange={(e, option: any): void => {
                          solDataHandler(
                            "newRequest",
                            "project",
                            option["text"]
                          );
                        }}
                        selectedKey={solNewRequestResponseData.project}
                      />
                    </div>
                    <div>
                      <Label
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalBoxReadOnlyLabelStyles
                          //   :
                          solModalBoxLabelStyles
                        }
                      >
                        Product
                      </Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalReadOnlyDropdownStyles
                          //   :
                          solNewRequestResponseData.productValidation
                            ? solModalErrorDropdownStyles
                            : solModalDropdownStyles
                        }
                        // disabled={
                        //   solPopup == "updateRequestPopup" ? true : false
                        // }
                        options={solAllDropDownOptions.ProductOptns}
                        dropdownWidth={"auto"}
                        onChange={(e, option: any): void => {
                          solDataHandler(
                            "newRequest",
                            "product",
                            option["text"]
                          );
                        }}
                        selectedKey={solNewRequestResponseData.product}
                      />
                    </div>
                  </div>
                  <div
                    style={{ display: "flex", justifyContent: "space-between" }}
                  >
                    <div>
                      <Label
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalBoxReadOnlyLabelStyles
                          //   :
                          solModalBoxLabelStyles
                        }
                      >
                        Subject
                      </Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalReadOnlyDropdownStyles
                          //   :
                          solNewRequestResponseData.subjectValidation
                            ? solModalErrorDropdownStyles
                            : solModalDropdownStyles
                        }
                        // disabled={
                        //   solPopup == "updateRequestPopup" ? true : false
                        // }
                        options={solAllDropDownOptions.SubjectOptns}
                        dropdownWidth={"auto"}
                        onChange={(e, option: any): void => {
                          solDataHandler(
                            "newRequest",
                            "subject",
                            option["text"]
                          );
                        }}
                        selectedKey={solNewRequestResponseData.subject}
                      />
                    </div>
                    <div>
                      <Label
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalBoxReadOnlyLabelStyles
                          //   :
                          solModalBoxLabelStyles
                        }
                      >
                        Location
                      </Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalReadOnlyDropdownStyles
                          //   :
                          solNewRequestResponseData.locationValidation
                            ? solModalErrorDropdownStyles
                            : solModalDropdownStyles
                        }
                        multiSelect
                        // disabled={
                        //   solPopup == "updateRequestPopup" ? true : false
                        // }
                        options={solAllDropDownOptions.LocationOptns}
                        dropdownWidth={"auto"}
                        onChange={(e, option: any): void => {
                          solDataHandler("newRequest", "location", option);
                        }}
                        selectedKeys={
                          solNewRequestResponseData.location.length > 0
                            ? solNewRequestResponseData.location
                            : []
                        }
                      />
                    </div>
                    <div>
                      <Label
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalBoxReadOnlyLabelStyles
                          //   :
                          solModalBoxLabelStyles
                        }
                      >
                        Name of stock
                      </Label>
                      <TextField
                        styles={
                          // solPopup == "updateRequestPopup"
                          //   ? solModalReadOnlyTxtBoxStyles
                          //   :
                          solNewRequestResponseData.stockNameValidation
                            ? solModalErrorTxtBoxStyles
                            : solModalTxtBoxStyles
                        }
                        // readOnly={
                        //   solPopup == "updateRequestPopup" ? true : false
                        // }
                        value={solNewRequestResponseData.stockName}
                        onChange={(e, value: string): void => {
                          solDataHandler("newRequest", "stockName", value);
                        }}
                      />
                    </div>
                  </div>
                  <div
                    style={{ display: "flex", justifyContent: "space-between" }}
                  >
                    <div>
                      <Label styles={solModalBoxLabelStyles}>
                        {solPopup == "newRequestPopup"
                          ? "Availability"
                          : solPopup == "updateRequestPopup"
                          ? "Add availability"
                          : null}
                      </Label>
                      <TextField
                        styles={
                          solNewRequestResponseData.availabilityValidation
                            ? solModalErrorTxtBoxStyles
                            : solModalTxtBoxStyles
                        }
                        value={solNewRequestResponseData.availability}
                        onChange={(e, value: string): void => {
                          /^[0-9]+$|^$/.test(value)
                            ? solDataHandler(
                                "newRequest",
                                "availability",
                                value
                              )
                            : "";
                        }}
                      />
                    </div>
                  </div>
                </div>
                <div className={styles.solModalBoxButtonSection}>
                  {solNewRequestResponseData.overAllValidation &&
                  solNewRequestResponseData.availability == "0" ? (
                    <Label style={{ color: "#f00", fontWeight: 600 }}>
                      * Invalid availability
                    </Label>
                  ) : solNewRequestResponseData.overAllValidation ? (
                    <Label style={{ color: "#f00", fontWeight: 600 }}>
                      * All fields are mandatory
                    </Label>
                  ) : solNewRequestResponseData.duplicateValidation ? (
                    <Label style={{ color: "#f00", fontWeight: 600 }}>
                      * The stock already exists
                    </Label>
                  ) : null}

                  <button
                    className={styles.solSubmitBtn}
                    onClick={(): void => {
                      setSolLoader("newRequestLoader");
                      validationFunction("newRequest");
                    }}
                  >
                    {solLoader == "newRequestLoader" ? (
                      <Spinner />
                    ) : (
                      <>
                        <Icon
                          iconName="Save"
                          style={{ position: "relative", top: 2, left: -8 }}
                        />
                        {solPopup == "newRequestPopup"
                          ? "Submit"
                          : solPopup == "updateRequestPopup"
                          ? "Update"
                          : null}
                      </>
                    )}
                  </button>
                  <button
                    className={styles.solCloseBtn}
                    onClick={(): void => {
                      solLoader == "newRequestLoader"
                        ? null
                        : (setSolNewRequestResponseData(solNewReqResponseData),
                          setSolPopup(""));
                    }}
                  >
                    <Icon
                      iconName="Cancel"
                      style={{ position: "relative", top: 2, left: -8 }}
                    />
                    Cancel
                  </button>
                </div>
              </Modal>
            ) : (
              ""
            )}
          </div>
          {/* New & Update Request-Popup Ends */}
          {/* Add Request-Popup Starts */}
          <div>
            {solPopup == "addRequestPopup" ? (
              <Modal
                isOpen={solPopup == "addRequestPopup" ? true : false}
                isBlocking={true}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    padding: "20px 0",
                    width: 600,
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
                    <Label className={styles.newRequestTitle}>
                      Add request
                    </Label>
                  </div>
                </div>
                <div style={{ width: "85%", margin: "auto" }}>
                  <div
                    style={{ display: "flex", justifyContent: "space-between" }}
                  >
                    <div>
                      <Label styles={solModalBoxLabelStyles}>Request for</Label>
                      <Label>{solAddRequestResponseData.stockName}</Label>
                    </div>
                    <div>
                      <Label styles={solModalDropdownStyles}>School</Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={solModalDropdownStyles}
                        options={solAllDropDownOptions.schoolOptns}
                        dropdownWidth={"auto"}
                        onChange={(e, option: any): void => {
                          solDataHandler(
                            "addRequest",
                            "school",
                            option["text"]
                          );
                        }}
                        selectedKey={solAddRequestResponseData.school}
                      />
                    </div>
                  </div>
                  <div
                    style={{ display: "flex", justifyContent: "space-between" }}
                  >
                    <div>
                      <Label required={true} styles={solModalBoxLabelStyles}>
                        Request by
                      </Label>
                      <NormalPeoplePicker
                        styles={
                          solAddRequestResponseData.requestFromValidation
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
                          return (
                            people.ID ==
                            (solAddRequestResponseData.requestFrom
                              ? solAddRequestResponseData.requestFrom
                              : null)
                          );
                        })}
                        onChange={(selectedUser): void => {
                          solDataHandler(
                            "addRequest",
                            "requestFrom",
                            selectedUser[0] ? selectedUser[0]["ID"] : null
                          );
                        }}
                      />
                    </div>
                    <div>
                      <Label
                        required={true}
                        style={{
                          marginTop: 0,
                        }}
                        styles={solModalBoxLabelStyles}
                      >
                        Count
                      </Label>
                      <TextField
                        styles={
                          solAddRequestResponseData.requestCountValidation
                            ? solModalErrorTxtBoxStyles
                            : solModalTxtBoxStyles
                        }
                        value={solAddRequestResponseData.requestCount}
                        onChange={(e, value: string): void => {
                          /^[0-9]+$|^$/.test(value)
                            ? solDataHandler(
                                "addRequest",
                                "requestCount",
                                value
                              )
                            : null;
                        }}
                      />
                    </div>
                  </div>
                </div>
                <div
                  style={{
                    width: "85%",
                  }}
                  className={styles.solModalBoxButtonSection}
                >
                  {solAddRequestResponseData.overAllValidation ? (
                    <Label style={{ color: "#f00", fontWeight: 600 }}>
                      * Fields are mandatory
                    </Label>
                  ) : null}
                  <button
                    className={styles.solSubmitBtn}
                    onClick={(): void => {
                      setSolLoader("addRequestLoader");
                      validationFunction("addRequest");
                    }}
                  >
                    {solLoader == "addRequestLoader" ? (
                      <Spinner />
                    ) : (
                      <>
                        <Icon
                          iconName="Save"
                          style={{ position: "relative", top: 2, left: -8 }}
                        />
                        Add
                      </>
                    )}
                  </button>
                  <button
                    className={styles.solCloseBtn}
                    onClick={(): void => {
                      solLoader == "addRequestLoader" ? null : setSolPopup("");
                    }}
                  >
                    <Icon
                      iconName="Cancel"
                      style={{ position: "relative", top: 2, left: -8 }}
                    />
                    Cancel
                  </button>
                </div>
              </Modal>
            ) : (
              ""
            )}
          </div>
          {/* Add Request-Popup Ends */}
          {/* Response List-Popup Starts */}
          <div>
            {solPopup == "responsesPopup" ? (
              <Modal
                isOpen={solPopup == "responsesPopup" ? true : false}
                isBlocking={true}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    width: 1150,
                    padding: "20px 0",
                    paddingBottom: "10px",
                    marginBottom: "10px",
                  }}
                >
                  <Label className={styles.newRequestTitle}>
                    Stock response list
                  </Label>
                </div>

                <div
                  style={{
                    width: "85%",
                    margin: "auto",
                    display: "flex",
                    justifyContent: "space-between",
                    marginBottom: "10px",
                  }}
                >
                  <Label>
                    Name of stock :{" "}
                    <span style={{ color: "#0882A5" }}>
                      {
                        solMasterData.filter((item) => {
                          return item.ID == solRequestData[0].stockID;
                        })[0].Stocks
                      }
                    </span>
                  </Label>
                  <Label>
                    Availability :{" "}
                    <span style={{ color: "#0882A5" }}>
                      {
                        solMasterData.filter((item) => {
                          return item.ID == solRequestData[0].stockID;
                        })[0].Availability
                      }
                    </span>
                  </Label>
                </div>

                <div
                  style={{
                    width: "85%",
                    margin: "auto",
                    display: "flex",
                    justifyContent: "space-between",
                    marginBottom: "10px",
                  }}
                >
                  <Pivot styles={solPivotStyles}>
                    <PivotItem
                      headerText="Requests"
                      headerButtonProps={{
                        "data-order": 1,
                        "data-title": "Requests",
                      }}
                    >
                      {/* <div
                        style={{
                          // width: "85%",
                          // margin: "auto",
                          display: "flex",
                          justifyContent: "space-between",
                          marginBottom: "10px",
                        }}
                      >
                        <Label>
                          Stock Name :{" "}
                          {
                            solMasterData.filter((item) => {
                              return item.ID == solRequestData[0].stockID;
                            })[0].Stocks
                          }
                        </Label>
                        <Label>
                          Availability :{" "}
                          {
                            solMasterData.filter((item) => {
                              return item.ID == solRequestData[0].stockID;
                            })[0].Availability
                          }
                        </Label>
                      </div> */}

                      <div style={{ width: "1000px" }}>
                        {/* DetailList-Section Starts */}
                        <div>
                          <DetailsList
                            items={solRequestData.filter((item) => {
                              return item.requestStatus == "Pending";
                            })}
                            columns={solRequestsColumns}
                            styles={
                              solRequestData.filter((item) => {
                                return item.requestStatus == "Pending";
                              }).length > 0
                                ? solModalDetailsListStyles
                                : solModalDetailsListNoDataStyles
                            }
                            setKey="set"
                            layoutMode={DetailsListLayoutMode.justified}
                            selectionMode={SelectionMode.none}
                          />
                        </div>
                        {/* DetailList-Section Ends */}
                        {solRequestData.filter((item) => {
                          return item.requestStatus == "Pending";
                        }).length > 0 ? null : (
                          <div
                            style={{
                              display: "flex",
                              justifyContent: "center",
                              alignItems: "center",
                              // marginTop: 15,
                              // marginBottom: 15,
                              height: "145px",
                            }}
                          >
                            <Label style={{ color: "#2392B2" }}>
                              No pending requests !!!
                            </Label>
                          </div>
                        )}
                      </div>

                      <div
                        style={{
                          // width: "85%",
                          // margin: "auto",
                          display: "flex",
                          justifyContent: "flex-end",
                        }}
                      >
                        {solRequestResponseData.filter((item) => {
                          return item.status == "Approved";
                        }).length > 0 ? (
                          <Label
                            style={
                              calculateProposedQty(solRequestResponseData) >
                              solMasterData.filter((item) => {
                                return item.ID == solRequestData[0].stockID;
                              })[0].Availability
                                ? { color: "#f00" }
                                : { color: "#000" }
                            }
                          >
                            Proposed Quantity :{" "}
                            {calculateProposedQty(solRequestResponseData)}
                          </Label>
                        ) : null}
                      </div>

                      <div
                        style={{
                          // width: "85%",
                          // margin: "auto",
                          display: "flex",
                          justifyContent: "space-between",
                          marginBottom: 10,
                          marginTop: 20,
                        }}
                      >
                        <div
                          style={{
                            display: "flex",
                            justifyContent: "space-between",
                            width: "50%",
                          }}
                        >
                          <Label>
                            Approved :{" "}
                            <span style={{ color: "#00a71c" }}>
                              {
                                solRequestData.filter((item) => {
                                  return item.requestStatus == "Approved";
                                }).length
                              }
                            </span>
                          </Label>
                          <Label>
                            Rejected :{" "}
                            <span style={{ color: "#ff0000" }}>
                              {
                                solRequestData.filter((item) => {
                                  return item.requestStatus == "Rejected";
                                }).length
                              }
                            </span>
                          </Label>
                          <Label>
                            Pending :{" "}
                            <span style={{ color: "#F46700" }}>
                              {
                                solRequestData.filter((item) => {
                                  return item.requestStatus == "Pending";
                                }).length
                              }
                            </span>
                          </Label>
                        </div>
                        <div className={styles.solResReqModalBoxButtonSection}>
                          {props.isAdmin ? (
                            <button
                              className={
                                solRequestResponseData.length == 0
                                  ? styles.solSubmitBtnDisabled
                                  : calculateProposedQty(
                                      solRequestResponseData
                                    ) >
                                    parseInt(
                                      solMasterData.filter((item) => {
                                        return (
                                          item.ID == solRequestData[0].stockID
                                        );
                                      })[0].Availability
                                    )
                                  ? styles.solSubmitBtnDisabled
                                  : styles.solSubmitBtn
                              }
                              onClick={(): void => {
                                solRequestResponseData.length == 0
                                  ? null
                                  : calculateProposedQty(
                                      solRequestResponseData
                                    ) >
                                    parseInt(
                                      solMasterData.filter((item) => {
                                        return (
                                          item.ID == solRequestData[0].stockID
                                        );
                                      })[0].Availability
                                    )
                                  ? null
                                  : (setSolLoader("submitRequestLoader"),
                                    submitRequestFunction(
                                      solRequestResponseData
                                    ));
                              }}
                            >
                              {solLoader == "submitRequestLoader" ? (
                                <Spinner />
                              ) : (
                                <>
                                  <Icon
                                    iconName="Save"
                                    style={{
                                      position: "relative",
                                      top: 2,
                                      left: -8,
                                    }}
                                  />
                                  Submit
                                </>
                              )}
                            </button>
                          ) : null}
                          <button
                            className={styles.solCloseBtn}
                            onClick={(): void => {
                              solLoader == "submitRequestLoader"
                                ? null
                                : (setSolPopup(""),
                                  setSolRequestResponseData([]));
                            }}
                          >
                            <Icon
                              iconName="Cancel"
                              style={{ position: "relative", top: 2, left: -8 }}
                            />
                            {props.isAdmin ? "Cancel" : "Close"}
                          </button>
                        </div>
                      </div>
                    </PivotItem>
                    <PivotItem
                      headerText="Responses"
                      headerButtonProps={{
                        "data-order": 2,
                        "data-title": "Responses",
                      }}
                    >
                      <div style={{ width: "1000px" }}>
                        {/* DetailList-Section Starts */}
                        <div>
                          <DetailsList
                            items={solRequestData.filter((item) => {
                              return (
                                item.requestStatus == "Approved" ||
                                item.requestStatus == "Rejected"
                              );
                            })}
                            columns={solRequestsColumns}
                            styles={
                              solRequestData.filter((item) => {
                                return (
                                  item.requestStatus == "Approved" ||
                                  item.requestStatus == "Rejected"
                                );
                              }).length > 0
                                ? solModalDetailsListStyles
                                : solModalDetailsListNoDataStyles
                            }
                            setKey="set"
                            layoutMode={DetailsListLayoutMode.justified}
                            selectionMode={SelectionMode.none}
                          />
                        </div>
                        {/* DetailList-Section Ends */}
                        {solRequestData.filter((item) => {
                          return (
                            item.requestStatus == "Approved" ||
                            item.requestStatus == "Rejected"
                          );
                        }).length > 0 ? null : (
                          <div
                            style={{
                              display: "flex",
                              justifyContent: "center",
                              alignItems: "center",
                              // marginTop: 15,
                              // marginBottom: 15,
                              height: "145px",
                            }}
                          >
                            <Label style={{ color: "#2392B2" }}>
                              No responses !!!
                            </Label>
                          </div>
                        )}
                      </div>

                      <div
                        style={{
                          // width: "85%",
                          // margin: "auto",
                          display: "flex",
                          justifyContent: "space-between",
                          marginBottom: 10,
                          marginTop: 20,
                        }}
                      >
                        <div
                          style={{
                            display: "flex",
                            justifyContent: "space-between",
                            width: "50%",
                          }}
                        >
                          <Label>
                            Approved :{" "}
                            <span style={{ color: "#00a71c" }}>
                              {
                                solRequestData.filter((item) => {
                                  return item.requestStatus == "Approved";
                                }).length
                              }
                            </span>
                          </Label>
                          <Label>
                            Rejected :{" "}
                            <span style={{ color: "#ff0000" }}>
                              {
                                solRequestData.filter((item) => {
                                  return item.requestStatus == "Rejected";
                                }).length
                              }
                            </span>
                          </Label>
                          <Label>
                            Pending :{" "}
                            <span style={{ color: "#F46700" }}>
                              {
                                solRequestData.filter((item) => {
                                  return item.requestStatus == "Pending";
                                }).length
                              }
                            </span>
                          </Label>
                        </div>
                        <div className={styles.solResReqModalBoxButtonSection}>
                          <button
                            className={styles.solCloseBtn}
                            onClick={(): void => {
                              solLoader == "submitRequestLoader"
                                ? null
                                : (setSolPopup(""),
                                  setSolRequestResponseData([]));
                            }}
                          >
                            <Icon
                              iconName="Cancel"
                              style={{ position: "relative", top: 2, left: -8 }}
                            />
                            {props.isAdmin ? "Cancel" : "Close"}
                          </button>
                        </div>
                      </div>
                    </PivotItem>
                  </Pivot>
                </div>
              </Modal>
            ) : (
              ""
            )}
          </div>
          {/* Response List-Popup Ends */}
        </div>
        {/* Modal-Section Ends */}

        <div>
          <Modal isOpen={solDeletePopup.condition} isBlocking={true}>
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
                <Label className={styles.deletePopupTitle}>Delete stock</Label>
                <Label
                  style={{
                    padding: "5px 20px",
                  }}
                  className={styles.deletePopupDesc}
                >
                  Are you sure want to delete?
                </Label>
              </div>
            </div>
            <div className={styles.apDeletePopupBtnSection}>
              <button
                onClick={(_) => {
                  setSolLoader("dltLoader");
                  solDeleteItem(solDeletePopup.targetId);
                }}
                className={styles.apDeletePopupYesBtn}
              >
                {solLoader == "dltLoader" ? <Spinner /> : "Yes"}
              </button>
              <button
                onClick={(_) => {
                  setsolDeletePopup({ condition: false, targetId: 0 });
                }}
                className={styles.apDeletePopupNoBtn}
              >
                No
              </button>
            </div>
          </Modal>
        </div>
      </div>
    </>
  );
};

export default StockList;
