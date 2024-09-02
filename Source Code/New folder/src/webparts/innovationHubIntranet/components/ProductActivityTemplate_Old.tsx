import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
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
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";
import { arraysEqual, IDetailsListStyles } from "office-ui-fabric-react";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

let sortPATData = [];
let sortPATFilter = [];

const ProductActivityTemplate = (props: any) => {
  const sharepointWeb = Web(props.URL);

  let CurrentPage = 1;
  let totalPageItems = 10;

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  const _patColumns = [
    // {
    //   key: "Column1",
    //   name: "Type",
    //   fieldName: "Types",
    //   minWidth: 150,
    //   maxWidth: 150,
    //   onRender: (item) => (
    //     <>
    //       <TooltipHost
    //         id={item.ID}
    //         content={item.Types}
    //         overflowMode={TooltipOverflowMode.Parent}
    //       >
    //         <span aria-describedby={item.ID}>{item.Types}</span>
    //       </TooltipHost>
    //     </>
    //   ),
    // },
    // {
    //   key: "Column2",
    //   name: "Area/Stream",
    //   fieldName: "Area",
    //   minWidth: 230,
    //   maxWidth: 230,
    //   onRender: (item) => (
    //     <>
    //       <TooltipHost
    //         id={item.ID}
    //         content={item.Area}
    //         overflowMode={TooltipOverflowMode.Parent}
    //       >
    //         <span aria-describedby={item.ID}>{item.Area}</span>
    //       </TooltipHost>
    //     </>
    //   ),
    // },
    {
      key: "Column3",
      name: "Product(Program)",
      fieldName: "Product",
      minWidth: 280,
      maxWidth: 380,
      onColumnClick: (ev, column) => {
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
      key: "Column4",
      name: "Project",
      fieldName: "Project",
      minWidth: 280,
      maxWidth: 380,
      onColumnClick: (ev, column) => {
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
      key: "Column5",
      name: "Code",
      fieldName: "Code",
      minWidth: 80,
      maxWidth: 180,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column6",
      name: "Template",
      fieldName: "Title",
      minWidth: 200,
      maxWidth: 400,
      onColumnClick: (ev, column) => {
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
      key: "Column7",
      name: "Completion",
      fieldName: "Completion",
      minWidth: 220,
      maxWidth: 400,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <div style={{ width: 200, position: "relative" }}>
          {item.Completion >= 75 ? (
            <div className={PATstatusStyleClass.default}>
              <span className={PATstatusStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PATstatusStyleClass.completed}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : item.Completion >= 50 ? (
            <div className={PATstatusStyleClass.default}>
              <span className={PATstatusStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PATstatusStyleClass.scheduled}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : item.Completion >= 25 ? (
            <div className={PATstatusStyleClass.default}>
              <span className={PATstatusStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PATstatusStyleClass.onSchedule}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : item.Completion >= 0 ? (
            <div className={PATstatusStyleClass.default}>
              <span className={PATstatusStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PATstatusStyleClass.behindScheduled}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : (
            ""
          )}
        </div>
      ),
    },
    {
      key: "Column8",
      name: "Link",
      fieldName: "Link",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item) => (
        <>
          <Icon
            iconName="Link12"
            className={PATIconStyleClass.link}
            onClick={() => {
              props.handleclick("ProductActivityPlan", item.Id);
            }}
          />
        </>
      ),
    },
  ];
  const PATDrpDwnOptns = {
    Project: [{ key: "All", text: "All" }],
    Code: [{ key: "All", text: "All" }],
    Product: [{ key: "All", text: "All" }],
  };
  const PATFilterKeys = {
    Project: "All",
    Product: "All",
    Code: "",
    Template: "",
  };
  const PATstatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    height: 17,
  });
  const PATstatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      PATstatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      PATstatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      PATstatusStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      PATstatusStyle,
    ],
    default: [
      {
        fontWeight: "600",
        position: "relative",
        backgroundColor: "#edebe9",
      },
      PATstatusStyle,
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
  const PATdropdownStyles: Partial<IDropdownStyles> = {
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
  const PATActivedropdownStyles: Partial<IDropdownStyles> = {
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
  const PATSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 200,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
      marginTop: "3px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const PATActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 200,
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
  const PATlabelStyles = mergeStyleSets({
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
  const ATIconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const PATIconStyleClass = mergeStyleSets({
    link: [{ color: "#2392B2", margin: "0" }, ATIconStyle],
    delete: [{ color: "#CB1E06", margin: "0 7px " }, ATIconStyle],
    edit: [{ color: "#2392B2", margin: "0 7px 0 0" }, ATIconStyle],
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 20,
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
  const getActivityTemplate = () => {
    let _PATdata = [];
    sharepointWeb.lists
      .getByTitle("Activity Plan")
      .items.top(5000)
      .get()
      .then(async (items) => {
        items.forEach((item) => {
          _PATdata.push({
            Id: item.Id,
            Title: item.Title,
            Project: item.Project,
            Types: item.Types,
            Area: item.Area,
            Product: item.Product,
            Code: item.ProductCode,
            Completion: item.Completion ? item.Completion : 0,
          });
        });

        setPATData([..._PATdata]);
        sortPATData = _PATdata;
        sortPATFilter = _PATdata;
        setPATMasterData([..._PATdata]);
        paginateFunction(1, [..._PATdata]);
        reloadFilterOptions([..._PATdata]);
        setPATLoader("noLoader");
      })
      .catch((err) => {
        PATErrorFunction(err, "getActivityTemplate");
      });
  };
  const totalCompletion = (data) => {
    var sum: number = 0;
    if (data.length > 0) {
      data.forEach((x) => {
        sum += parseInt(x.Completion ? x.Completion : 0);
      });
      let avg = sum / parseInt(data.length);
      return avg ? Math.round(avg) : 0;
    } else {
      return 0;
    }
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;

    tempArrReload.forEach((at) => {
      if (
        PATDrpDwnOptns.Project.findIndex((prj) => {
          return prj.key == at.Project;
        }) == -1 &&
        at.Project
      ) {
        PATDrpDwnOptns.Project.push({
          key: at.Project,
          text: at.Project,
        });
      }
      if (
        PATDrpDwnOptns.Code.findIndex((cd) => {
          return cd.key == at.Code;
        }) == -1 &&
        at.Code
      ) {
        PATDrpDwnOptns.Code.push({
          key: at.Code,
          text: at.Code,
        });
      }
      if (
        PATDrpDwnOptns.Product.findIndex((prd) => {
          return prd.key == at.Product;
        }) == -1 &&
        at.Product
      ) {
        PATDrpDwnOptns.Product.push({
          key: at.Product,
          text: at.Product,
        });
      }
    });
    setPATDropDownOptions(PATDrpDwnOptns);
  };
  const PATListFilter = (key, option) => {
    let tempArr = [...PATData];
    let tempDpFilterKeys = { ...PATFilterOptions };
    tempDpFilterKeys[`${key}`] = option;

    if (tempDpFilterKeys.Project != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Project == tempDpFilterKeys.Project;
      });
    }
    if (tempDpFilterKeys.Product != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Product == tempDpFilterKeys.Product;
      });
    }
    if (tempDpFilterKeys.Code) {
      tempArr = tempArr.filter((arr) => {
        return arr.Code
          ? arr.Code.toLowerCase().includes(tempDpFilterKeys.Code.toLowerCase())
          : "";
      });
    }
    if (tempDpFilterKeys.Template) {
      tempArr = tempArr.filter((arr) => {
        return arr.Title
          ? arr.Title.toLowerCase().includes(
              tempDpFilterKeys.Template.toLowerCase()
            )
          : "";
      });
    }
    sortPATFilter = tempArr;
    paginateFunction(1, [...tempArr]);
    setPATFilterOptions({ ...tempDpFilterKeys });
  };
  const paginateFunction = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      CurrentPage = pagenumber;
      setPATDisplayData(paginatedItems);
      setPATCurrentPage(pagenumber);
    } else {
      setPATDisplayData([]);
      setPATCurrentPage(1);
    }
  };
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const PATErrorFunction = (error: any, functionName: string) => {
    console.log(error);

    let response = {
      ComponentName: "Product activity template",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setPATLoader("noLoader");
        ErrorPopup();
      }
    );
  };
  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _patColumns;
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

    const newPATData = _copyAndSort(
      sortPATData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newPATFilter = _copyAndSort(
      sortPATFilter,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setPATData([...newPATData]);
    sortPATFilter = newPATFilter;
    paginateFunction(1, newPATFilter);
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

  // Use State
  const [PATReRender, setPATReRender] = useState(false);
  const [PATMasterData, setPATMasterData] = useState([]);
  const [PATData, setPATData] = useState([]);
  const [PATDisplayData, setPATDisplayData] = useState([]);
  const [PATDropDownOptions, setPATDropDownOptions] = useState(PATDrpDwnOptns);
  const [PATFilterOptions, setPATFilterOptions] = useState(PATFilterKeys);
  const [PATcurrentPage, setPATCurrentPage] = useState(CurrentPage);
  const [PATLoader, setPATLoader] = useState("noLoader");
  const [patColumns, setpatColumns] = useState(_patColumns);

  //Use Effect
  useEffect(() => {
    setPATLoader("startUpLoader");
    getActivityTemplate();
  }, [PATReRender]);

  return (
    <div style={{ padding: "5px 15px" }}>
      {PATLoader == "startUpLoader" ? <CustomLoader /> : null}
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
          <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
            Product List
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
            <Label className={PATlabelStyles.inputLabels}>
              Product(Program)
            </Label>
            <Dropdown
              dropdownWidth="auto"
              selectedKey={PATFilterOptions.Product}
              placeholder="Select an option"
              options={PATDropDownOptions.Product}
              styles={
                PATFilterOptions.Product != "All"
                  ? PATActivedropdownStyles
                  : PATdropdownStyles
              }
              onChange={(e, option: any) => {
                PATListFilter("Product", option["key"]);
              }}
            />
          </div>
          <div>
            <Label className={PATlabelStyles.inputLabels}>Project</Label>
            <Dropdown
              selectedKey={PATFilterOptions.Project}
              placeholder="Select an option"
              options={PATDropDownOptions.Project}
              styles={
                PATFilterOptions.Project != "All"
                  ? PATActivedropdownStyles
                  : PATdropdownStyles
              }
              onChange={(e, option: any) => {
                PATListFilter("Project", option["key"]);
              }}
            />
          </div>
          <div>
            <Label className={PATlabelStyles.inputLabels}>Code</Label>
            <SearchBox
              styles={
                PATFilterOptions.Code
                  ? PATActiveSearchBoxStyles
                  : PATSearchBoxStyles
              }
              value={PATFilterOptions.Code}
              onChange={(e, value) => {
                PATListFilter("Code", value);
              }}
            />
          </div>
          <div>
            <Label className={PATlabelStyles.inputLabels}>Template</Label>
            <SearchBox
              styles={
                PATFilterOptions.Template
                  ? PATActiveSearchBoxStyles
                  : PATSearchBoxStyles
              }
              value={PATFilterOptions.Template}
              onChange={(e, value) => {
                PATListFilter("Template", value);
              }}
            />
          </div>
          <div>
            <Icon
              iconName="Refresh"
              title="Click to reset"
              className={PATIconStyleClass.refresh}
              onClick={() => {
                setPATFilterOptions({ ...PATFilterKeys });
                paginateFunction(1, [...PATMasterData]);
                setPATData([...PATMasterData]);
                sortPATData = PATMasterData;
                sortPATFilter = PATMasterData;
                setpatColumns(_patColumns);
              }}
            />
          </div>
        </div>
        <div>
          <Label style={{ marginRight: 5 }}>
            Number of records :{" "}
            <span style={{ color: "#038387" }}>{PATDisplayData.length}</span>
          </Label>
        </div>
      </div>
      <div>
        <DetailsList
          items={PATDisplayData}
          columns={patColumns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          // styles={gridStyles}
          styles={{ root: { width: "100%" } }}
        />
      </div>
      {PATDisplayData.length > 0 ? (
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            margin: "10px 0",
          }}
        >
          <Pagination
            currentPage={PATcurrentPage}
            totalPages={
              PATData.length > 0
                ? Math.ceil(PATData.length / totalPageItems)
                : 1
            }
            onChange={(page) => {
              paginateFunction(page, PATData);
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
export default ProductActivityTemplate;
