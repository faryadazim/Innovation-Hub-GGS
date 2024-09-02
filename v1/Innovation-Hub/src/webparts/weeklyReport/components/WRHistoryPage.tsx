import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import {
  DetailsList,
  IDetailsListStyles,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  TooltipHost,
  TooltipOverflowMode,
  IColumn,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./WeeklyReport.module.scss";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

let sortData = [];
let sortFilterData = [];

interface IHistoryData {
  condition: boolean;
  sourcePage: string;
  targetID: number;
}
interface IProps {
  context: WebPartContext;
  spcontext: any;
  graphContent: any;
  URL: string;
  peopleList: any;
  historyDataHandler: any;
  historyData: IHistoryData;
}

const DocumentReviewHistory = (props: IProps) => {
  const sharepointWeb = Web(props.URL);
  const DocumentID = props.historyData.targetID;
  const PageName = props.historyData.sourcePage;

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  const _historyColumns = [
    {
      key: "Column1",
      name: "File",
      fieldName: "FileName",
      minWidth: 150,
      maxWidth: 250,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.FileName}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.FileName}</span>
        </TooltipHost>
      ),
    },
    {
      key: "Column2",
      name: "Sent",
      fieldName: "Sent",
      minWidth: 80,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => moment(item.Sent).format("DD/MM/YYYY"),
    },
    {
      key: "Column3",
      name: "Request",
      fieldName: "Request",
      minWidth: 50,
      maxWidth: 80,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column4",
      name: "Response",
      fieldName: "Response",
      minWidth: 50,
      maxWidth: 80,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column5",
      name: "From user",
      fieldName: "From",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column6",
      name: "To user",
      fieldName: "To",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column7",
      name: "Document type",
      fieldName: "DocType",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    // {
    //   key: "Column8",
    //   name:
    //     PageName == "ReviewReport" ? "Request comments" : "Response comments",
    //   fieldName:
    //     PageName == "ReviewReport" ? "RequestComments" : "ResponseComments",
    //   minWidth: 200,
    //   maxWidth: 500,
    //   onColumnClick: (ev, column) => {
    //     _onColumnClick(ev, column);
    //   },
    //   onRender: (item) => (
    //     <TooltipHost
    //       id={item.ID}
    //       content={
    //         PageName == "ReviewReport"
    //           ? item.RequestComments
    //           : item.ResponseComments
    //       }
    //       overflowMode={TooltipOverflowMode.Parent}
    //     >
    //       <span aria-describedby={item.ID}>
    //         {PageName == "ReviewReport"
    //           ? item.RequestComments
    //           : item.ResponseComments}
    //       </span>
    //     </TooltipHost>
    //   ),
    // },
    {
      key: "Column8",
      name: "Request comments",
      fieldName: "RequestComments",
      minWidth: 200,
      maxWidth: 250,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.RequestComments}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.RequestComments}</span>
        </TooltipHost>
      ),
    },
    {
      key: "Column9",
      name: "Response comments",
      fieldName: "ResponseComments",
      minWidth: 200,
      maxWidth: 250,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.ResponseComments}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.ResponseComments}</span>
        </TooltipHost>
      ),
    },
  ];
  const historyDrpDwnOptns = {
    Request: [{ key: "All", text: "All" }],
    Response: [{ key: "All", text: "All" }],
    DocType: [{ key: "All", text: "All" }],
  };
  const _historyFilterKeys = {
    Request: "All",
    Response: "All",
    DocType: "All",
  };

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
      ".ms-DetailsHeader-cellTitle": {
        background: "#03828711 !important",
        color: "#038387 !important",
      },
    },
    headerWrapper: {
      flex: "0 0 auto",
    },
    contentWrapper: {
      flex: "1 1 auto",
      overflowY: "auto",
      overflowX: "hidden",
    },
  };
  const historyIconStyleClass = mergeStyleSets({
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
        marginTop: 28,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
  });
  const historyLabelStyles = mergeStyleSets({
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
  const historyDropDownStyles: Partial<IDropdownStyles> = {
    root: { width: 200, marginRight: 15 },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      border: "1px solid #E8E8EA",
      fontWeight: 600,
    },
    dropdownItemsWrapper: { backgroundColor: "#F5F5F7", fontSize: 12 },
    dropdownItemSelected: { backgroundColor: "#DCDCDC", fontSize: 12 },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const historyActivedropdownStyles: Partial<IDropdownStyles> = {
    root: { width: 200, marginRight: 15 },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      border: "2px solid #038387",
      color: "#038387",
    },
    dropdownItemsWrapper: { backgroundColor: "#F5F5F7", fontSize: 12 },
    dropdownItemSelected: { backgroundColor: "#DCDCDC", fontSize: 12 },
    caretDown: { fontSize: 14, color: "#038387", fontWeight: 600 },
  };

  // Use State
  const [historyReRender, setHistoryReRender] = useState(false);
  const [historyMasterData, setHistoryMasterData] = useState([]);
  const [historyData, setHistoryData] = useState([]);
  const [historyFilterData, setHistoryFilterData] = useState([]);
  const [historyDropDownOptions, setHistoryDropDownOptions] =
    useState(historyDrpDwnOptns);
  const [historyFilterKeys, setHistoryFilterKeys] =
    useState(_historyFilterKeys);
  const [historyColumns, setHistoryColumns] = useState(_historyColumns);
  const [historyLoader, setHistoryLoader] = useState("noLoader");
  // Function to render

  const getHistoryData = () => {
    sharepointWeb.lists
      .getByTitle("Review Log")
      .items.getById(DocumentID)
      .get()
      .then((record) => {
        let DocumentAuditLink = record.auditDocLink;
        let _DRHdata = [];
        sharepointWeb.lists
          .getByTitle("Review Log")
          .items.filter("auditDocLink eq '" + DocumentAuditLink + "' ")
          .select(
            "*",
            "FromUser/Title",
            "FromUser/Id",
            "FromUser/EMail",
            "ToUser/Title",
            "ToUser/Id",
            "ToUser/EMail",
            "CcEmail/Title",
            "CcEmail/Id",
            "CcEmail/EMail"
          )
          .expand("FromUser", "CcEmail", "ToUser")
          .top(5000)
          .get()
          .then(async (items) => {
            items.forEach((item) => {
              _DRHdata.push({
                FileName: item.Title,
                Sent: item.auditSent,
                Request: item.auditRequestType,
                Response: item.auditResponseType,
                From: item.auditFrom ? item.auditFrom : "",
                To: item.auditTo ? item.auditTo : "",
                DocType: item.auditDocType ? item.auditDocType : "",
                RequestComments: item.auditComments,
                ResponseComments: item.Response_x0020_Comments
                  ? item.Response_x0020_Comments.replace(/<[^>]*>/g, "")
                  : "",
              });
              setHistoryFilterData([..._DRHdata]);
              sortFilterData = _DRHdata;
              setHistoryData([..._DRHdata]);
              sortData = _DRHdata;
              setHistoryMasterData([..._DRHdata]);
              reloadFilterOptions([..._DRHdata]);
              setHistoryLoader("noLoader");
            });
          })
          .catch((err) => {
            historyErrorFunction(err, "getHistoryData-DRData-history");
          });
      })
      .catch((err) => {
        historyErrorFunction(err, "getHistoryData-DRData");
      });
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;

    tempArrReload.forEach((at) => {
      if (
        historyDrpDwnOptns.Request.findIndex((prj) => {
          return prj.key == at.Request;
        }) == -1 &&
        at.Request
      ) {
        historyDrpDwnOptns.Request.push({
          key: at.Request,
          text: at.Request,
        });
      }
      if (
        historyDrpDwnOptns.Response.findIndex((cd) => {
          return cd.key == at.Response;
        }) == -1 &&
        at.Response
      ) {
        historyDrpDwnOptns.Response.push({
          key: at.Response,
          text: at.Response,
        });
      }
      if (
        historyDrpDwnOptns.DocType.findIndex((prd) => {
          return prd.key == at.DocType;
        }) == -1 &&
        at.DocType
      ) {
        historyDrpDwnOptns.DocType.push({
          key: at.DocType,
          text: at.DocType,
        });
      }
    });
    setHistoryDropDownOptions(historyDrpDwnOptns);
  };
  const historyListFilter = (key, option) => {
    let tempArr = [...historyData];
    let tempFilterKeys = { ...historyFilterKeys };
    tempFilterKeys[`${key}`] = option;

    if (tempFilterKeys.Request != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Request == tempFilterKeys.Request;
      });
    }
    if (tempFilterKeys.Response != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Response == tempFilterKeys.Response;
      });
    }
    if (tempFilterKeys.DocType != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.DocType == tempFilterKeys.DocType;
      });
    }
    setHistoryFilterData([...tempArr]);
    sortFilterData = tempArr;
    setHistoryFilterKeys({ ...tempFilterKeys });
  };

  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const historyErrorFunction = (error: any, functionName: string) => {
    console.log(error, functionName);

    let response = {
      ComponentName: "Weekly report - history page",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setHistoryLoader("noLoader");
        ErrorPopup();
      }
    );
  };

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempColumns = _historyColumns;
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

    const newData = _copyAndSort(
      sortData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newFilterData = _copyAndSort(
      sortFilterData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setHistoryData([...newData]);
    setHistoryFilterData([...newFilterData]);
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

  //Use Effect
  useEffect(() => {
    setHistoryLoader("startUpLoader");
    getHistoryData();
  }, [historyReRender]);

  return (
    <>
      <div style={{ padding: "5px 15px" }}>
        {historyLoader == "startUpLoader" ? <CustomLoader /> : null}
        <div
          style={{
            display: "flex",
            alignItems: "flex-start",
            justifyContent: "space-between",
            marginBottom: 10,
            color: "#2392b2",
          }}
        >
          {/* Header Start */}
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
          >
            <Icon
              iconName="NavigateBack"
              className={historyIconStyleClass.navArrow}
              onClick={() => {
                props.historyDataHandler(false, null);
              }}
            />
            <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
              Document review history
            </Label>
          </div>
        </div>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            paddingBottom: "10px",
            flexWrap: "wrap",
          }}
        >
          <div
            style={{ display: "flex", alignItems: "center", flexWrap: "wrap" }}
          >
            <div>
              <Label className={historyLabelStyles.inputLabels}>Request</Label>
              <Dropdown
                selectedKey={historyFilterKeys.Request}
                placeholder="Select an option"
                options={historyDropDownOptions.Request}
                styles={
                  historyFilterKeys.Request != "All"
                    ? historyActivedropdownStyles
                    : historyDropDownStyles
                }
                onChange={(e, option: any) => {
                  historyListFilter("Request", option["key"]);
                }}
              />
            </div>
            <div>
              <Label className={historyLabelStyles.inputLabels}>Response</Label>
              <Dropdown
                selectedKey={historyFilterKeys.Response}
                placeholder="Select an option"
                options={historyDropDownOptions.Response}
                styles={
                  historyFilterKeys.Response != "All"
                    ? historyActivedropdownStyles
                    : historyDropDownStyles
                }
                onChange={(e, option: any) => {
                  historyListFilter("Response", option["key"]);
                }}
              />
            </div>
            <div>
              <Label className={historyLabelStyles.inputLabels}>
                Document type
              </Label>
              <Dropdown
                selectedKey={historyFilterKeys.DocType}
                placeholder="Select an option"
                options={historyDropDownOptions.DocType}
                styles={
                  historyFilterKeys.DocType != "All"
                    ? historyActivedropdownStyles
                    : historyDropDownStyles
                }
                onChange={(e, option: any) => {
                  historyListFilter("DocType", option["key"]);
                }}
              />
            </div>
            <div>
              <Icon
                iconName="Refresh"
                title="Click to reset"
                className={historyIconStyleClass.refresh}
                onClick={() => {
                  setHistoryFilterKeys(_historyFilterKeys);
                  setHistoryFilterData([...historyMasterData]);
                  setHistoryData([...historyMasterData]);
                  sortData = historyMasterData;
                  sortFilterData = historyMasterData;
                  setHistoryColumns(_historyColumns);
                }}
              />
            </div>
          </div>
          <div>
            <Label style={{ marginRight: 5, marginTop: 25 }}>
              Number of records :{" "}
              <span style={{ color: "#038387" }}>
                {historyFilterData.length}
              </span>
            </Label>
          </div>
        </div>
        <DetailsList
          layoutMode={DetailsListLayoutMode.justified}
          items={historyFilterData}
          columns={historyColumns}
          styles={gridStyles}
          setKey="set"
          selectionMode={SelectionMode.none}
        />
        {historyFilterData.length == 0 ? (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              marginTop: "15px",
            }}
          >
            <Label style={{ color: "#2392B2" }}>
              No data found !!!
              {/* This module under development!!! */}
            </Label>
          </div>
        ) : null}
      </div>
    </>
  );
};
export default DocumentReviewHistory;
