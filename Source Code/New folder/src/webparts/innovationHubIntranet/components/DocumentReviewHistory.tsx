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
import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

import Service from "../components/Services";

let sortDRHData = [];
let sortDRHFilter = [];

const DocumentReviewHistory = (props: any) => {
  const sharepointWeb = Web(props.URL);
  const DocumentID = props.DRID;

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  const _drhColumns = [
    {
      key: "Column1",
      name: "File",
      fieldName: "FileName",
      minWidth: 200,
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
      maxWidth: 80,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => moment(item.Sent).format("DD/MM/YYYY"),
    },
    {
      key: "Column3",
      name: "Request",
      fieldName: "Request",
      minWidth: 80,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column4",
      name: "Response",
      fieldName: "Response",
      minWidth: 80,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column5",
      name: "From user",
      fieldName: "From",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column6",
      name: "To user",
      fieldName: "To",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column7",
      name: "Document type",
      fieldName: "DocType",
      minWidth: 120,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
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
  const DRHDrpDwnOptns = {
    Request: [{ key: "All", text: "All" }],
    Response: [{ key: "All", text: "All" }],
    DocType: [{ key: "All", text: "All" }],
  };
  const DRHFilterKeys = {
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
  const DRGIconStyleClass = mergeStyleSets({
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
  const DRHlabelStyles = mergeStyleSets({
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
  const DRHdropdownStyles: Partial<IDropdownStyles> = {
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
  const DRHActivedropdownStyles: Partial<IDropdownStyles> = {
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
  const [DRHReRender, setDRHReRender] = useState(false);
  const [DRHMaster, setDRHMaster] = useState([]);
  const [DRHData, setDRHData] = useState([]);
  const [DRHFilter, setDRHFilter] = useState([]);
  const [DRHDropDownOptions, setDRHDropDownOptions] = useState(DRHDrpDwnOptns);
  const [DRHFilterOptions, setDRHFilterOptions] = useState(DRHFilterKeys);
  const [DRHLoader, setDRHLoader] = useState("noLoader");
  const [drhColumns, setdrhColumns] = useState(_drhColumns);
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
              setDRHFilter([..._DRHdata]);
              sortDRHFilter = _DRHdata;
              setDRHData([..._DRHdata]);
              sortDRHData = _DRHdata;
              setDRHMaster([..._DRHdata]);
              reloadFilterOptions([..._DRHdata]);
              setDRHLoader("noLoader");
            });
          })
          .catch((err) => {
            DRHErrorFunction(err, "getHistoryData-getDRData-auditDocLink");
          });
      })
      .catch((err) => {
        DRHErrorFunction(err, "getHistoryData-getDRData");
      });
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;

    tempArrReload.forEach((at) => {
      if (
        DRHDrpDwnOptns.Request.findIndex((prj) => {
          return prj.key == at.Request;
        }) == -1 &&
        at.Request
      ) {
        DRHDrpDwnOptns.Request.push({
          key: at.Request,
          text: at.Request,
        });
      }
      if (
        DRHDrpDwnOptns.Response.findIndex((cd) => {
          return cd.key == at.Response;
        }) == -1 &&
        at.Response
      ) {
        DRHDrpDwnOptns.Response.push({
          key: at.Response,
          text: at.Response,
        });
      }
      if (
        DRHDrpDwnOptns.DocType.findIndex((prd) => {
          return prd.key == at.DocType;
        }) == -1 &&
        at.DocType
      ) {
        DRHDrpDwnOptns.DocType.push({
          key: at.DocType,
          text: at.DocType,
        });
      }
    });
    setDRHDropDownOptions(DRHDrpDwnOptns);
  };
  const DRHListFilter = (key, option) => {
    let tempArr = [...DRHData];
    let tempDpFilterKeys = { ...DRHFilterOptions };
    tempDpFilterKeys[`${key}`] = option;

    if (tempDpFilterKeys.Request != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Request == tempDpFilterKeys.Request;
      });
    }
    if (tempDpFilterKeys.Response != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Response == tempDpFilterKeys.Response;
      });
    }
    if (tempDpFilterKeys.DocType != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.DocType == tempDpFilterKeys.DocType;
      });
    }
    setDRHFilter([...tempArr]);
    sortDRHFilter = tempArr;
    setDRHFilterOptions({ ...tempDpFilterKeys });
  };

  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const DRHErrorFunction = (error: any, functionName: string) => {
    console.log(error);
    let response = {
      ComponentName: "Review log history",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setDRHLoader("noLoader");
        ErrorPopup();
      }
    );
  };

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _drhColumns;
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

    const newDRHData = _copyAndSort(
      sortDRHData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newDRHFilter = _copyAndSort(
      sortDRHFilter,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setDRHData([...newDRHData]);
    setDRHFilter([...newDRHFilter]);
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
    setDRHLoader("startUpLoader");
    getHistoryData();
  }, [DRHReRender]);

  return (
    <>
      <div style={{ padding: "5px 15px" }}>
        {DRHLoader == "startUpLoader" ? <CustomLoader /> : null}
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
          <div className={styles.dpTitle}>
            <Icon
              iconName="NavigateBack"
              className={DRGIconStyleClass.navArrow}
              onClick={() => {
                props.handleclick("DocumentReview", null);
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
          <div className={styles.ddSection}>
            <div>
              <Label className={DRHlabelStyles.inputLabels}>Request</Label>
              <Dropdown
                selectedKey={DRHFilterOptions.Request}
                placeholder="Select an option"
                options={DRHDropDownOptions.Request}
                styles={
                  DRHFilterOptions.Request != "All"
                    ? DRHActivedropdownStyles
                    : DRHdropdownStyles
                }
                onChange={(e, option: any) => {
                  DRHListFilter("Request", option["key"]);
                }}
              />
            </div>
            <div>
              <Label className={DRHlabelStyles.inputLabels}>Response</Label>
              <Dropdown
                selectedKey={DRHFilterOptions.Response}
                placeholder="Select an option"
                options={DRHDropDownOptions.Response}
                styles={
                  DRHFilterOptions.Response != "All"
                    ? DRHActivedropdownStyles
                    : DRHdropdownStyles
                }
                onChange={(e, option: any) => {
                  DRHListFilter("Response", option["key"]);
                }}
              />
            </div>
            <div>
              <Label className={DRHlabelStyles.inputLabels}>
                Document type
              </Label>
              <Dropdown
                selectedKey={DRHFilterOptions.DocType}
                placeholder="Select an option"
                options={DRHDropDownOptions.DocType}
                styles={
                  DRHFilterOptions.DocType != "All"
                    ? DRHActivedropdownStyles
                    : DRHdropdownStyles
                }
                onChange={(e, option: any) => {
                  DRHListFilter("DocType", option["key"]);
                }}
              />
            </div>
            <div>
              <Icon
                iconName="Refresh"
                title="Click to reset"
                className={DRGIconStyleClass.refresh}
                onClick={() => {
                  setDRHFilterOptions(DRHFilterKeys);
                  setDRHFilter([...DRHMaster]);
                  setDRHData([...DRHMaster]);
                  sortDRHData = DRHMaster;
                  sortDRHFilter = DRHMaster;
                  setdrhColumns(_drhColumns);
                }}
              />
            </div>
          </div>
          <div>
            <Label style={{ marginRight: 5, marginTop: 25 }}>
              Number of records :{" "}
              <span style={{ color: "#038387" }}>{DRHFilter.length}</span>
            </Label>
          </div>
        </div>
        <DetailsList
          layoutMode={DetailsListLayoutMode.justified}
          items={DRHFilter}
          columns={drhColumns}
          styles={gridStyles}
          setKey="set"
          selectionMode={SelectionMode.none}
        />
        {DRHFilter.length == 0 ? (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              marginTop: "15px",
            }}
          >
            <Label style={{ color: "#2392B2" }}>
              No data Found !!!
              {/* This module under development!!! */}
            </Label>
          </div>
        ) : null}
      </div>
    </>
  );
};
export default DocumentReviewHistory;
