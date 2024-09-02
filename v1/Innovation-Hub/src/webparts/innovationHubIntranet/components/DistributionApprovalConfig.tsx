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
  SelectionMode,
  Icon,
  Label,
  ILabelStyles,
  NormalPeoplePicker,
  SearchBox,
  ISearchBoxStyles,
  Dropdown,
  IDropdownStyles,
  Modal,
  IColumn,
  Spinner,
  Toggle,
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
import { IDropdown } from "office-ui-fabric-react";

interface IProps {
  context: any;
  spcontext: any;
  graphContent: any;
  URL: string;
  handleclick: any;
  pageType: string;
  peopleList: IPeoplelist[];
  isAdmin: boolean;
}
interface IData {
  ID: number;

  ToId: number;
  ToName: string;
  ToEmail: string;

  BusinessArea: string;

  EmployeesChoices: string;
  ContractorsChoices: string;
  ManagementChoices: string;
  BoardChoices: string;
  SchoolsChoices: string;
  PrincipalsChoices: string;
  InstructionCoachesChoices: string;
  TeachersChoices: string;
  TeachingAssistantsChoices: string;
  SchoolTeamsChoices: string;
  ParentsChoices: string;
  CommunitiesChoices: string;
  OtherChoices: string;
  ProfessionalsChoices: string;
  ServicesChoices: string;
  FundingChoices: string;
  ResourcesChoices: string;
  GovernmentsChoices: string;
  PoliticiansChoices: string;
  IndigenousChoices: string;
  ProfessionalAssociationsChoices: string;
  MediaChoices: string;

  EmployeesPosition: string[];
  ContractorsPosition: string[];
  ManagementPosition: string[];
  BoardPosition: string[];
  SchoolsPosition: string[];
  PrincipalsPosition: string[];
  InstructionCoachesPosition: string[];
  TeachersPosition: string[];
  TeachingAssistantsPosition: string[];
  SchoolTeamsPosition: string[];
  ParentsPosition: string[];
  CommunitiesPosition: string[];
  OtherPosition: string[];
  ProfessionalsPosition: string[];
  ServicesPosition: string[];
  FundingPosition: string[];
  ResourcesPosition: string[];
  GovernmentsPosition: string[];
  PoliticiansPosition: string[];
  IndigenousPosition: string[];
  ProfessionalAssociationsPosition: string[];
  MediaPosition: string[];

  EmployeesPeople: IPeoplelist[];
  ContractorsPeople: IPeoplelist[];
  ManagementPeople: IPeoplelist[];
  BoardPeople: IPeoplelist[];
  SchoolsPeople: IPeoplelist[];
  PrincipalsPeople: IPeoplelist[];
  InstructionCoachesPeople: IPeoplelist[];
  TeachersPeople: IPeoplelist[];
  TeachingAssistantsPeople: IPeoplelist[];
  SchoolTeamsPeople: IPeoplelist[];
  ParentsPeople: IPeoplelist[];
  CommunitiesPeople: IPeoplelist[];
  OtherPeople: IPeoplelist[];
  ProfessionalsPeople: IPeoplelist[];
  ServicesPeople: IPeoplelist[];
  FundingPeople: IPeoplelist[];
  ResourcesPeople: IPeoplelist[];
  GovernmentsPeople: IPeoplelist[];
  PoliticiansPeople: IPeoplelist[];
  IndigenousPeople: IPeoplelist[];
  ProfessionalAssociationsPeople: IPeoplelist[];
  MediaPeople: IPeoplelist[];
}
interface IDacDropDown {
  businessAreaFilter: IDropdownOption[];

  allPosition: IDropdownOption[];
  allChoices: IDropdownOption[];
}
interface IDacFilterKeys {
  businessArea: string;
  toUser: string;
}

let columnSortArr: IData[] = [];
let columnSortMasterArr: IData[] = [];

const DistributionApprovalConfig = (props: IProps): JSX.Element => {
  // Variable-Declaration-Section Starts
  const sharepointWeb: IWeb = Web(props.URL);
  const DLApprovalConfigListName: string = "Distribution Approval Config";
  const MasterUserListName: string = "Master User List";
  const allPeoples: IPeoplelist[] = props.peopleList;

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  const dacColumns: IColumn[] = [
    {
      key: "column1",
      name: "Business area",
      fieldName: "BusinessArea",
      minWidth: 300,
      maxWidth: 500,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "column2",
      name: "Approver",
      fieldName: "ToName",
      minWidth: 300,
      maxWidth: 500,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },

    {
      key: "column3",
      name: "Action",
      fieldName: "",
      minWidth: 80,
      maxWidth: 100,

      onRender: (item: IData) => (
        <>
          <Icon
            iconName="Edit"
            title="Edit"
            className={dacIconStyleClass.edit}
            onClick={() => {
              setDacPopup({
                condition: true,
                responseData: { ...item },
                validation: false,
              });
            }}
          />
        </>
      ),
    },
  ];
  const dacDrpDwnOptns: IDacDropDown = {
    businessAreaFilter: [{ key: "All", text: "All" }],
    allPosition: [],
    allChoices: [
      // { key: "Not Mandatory", text: "Not Mandatory" },
      // { key: "Mandatory", text: "Mandatory" },
      { key: "No", text: "No" },
      { key: "Yes", text: "Yes" },
    ],
  };
  const dacFilterKeys: IDacFilterKeys = {
    businessArea: "All",
    toUser: "",
  };

  let currentpage: number = 1;
  let totalPageItems: number = 10;
  // Variable-Declaration-Section Ends
  // Styles-Section Starts
  const dacLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 165,
      marginTop: 5,
      marginRight: 10,
      fontSize: 13,
      color: "#323130",
    },
  };
  const dacSearchBoxStyles: Partial<ISearchBoxStyles> = {
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
  const dacActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
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
  const dacDropdownStyles: Partial<IDropdownStyles> = {
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
  const dacActiveDropdownStyles: Partial<IDropdownStyles> = {
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
  const dacModalDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      marginTop: 5,
      marginRight: 15,
      backgroundColor: "#fff",
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      color: "#000",
      border: "1px solid #000",
    },
    callout: {
      maxHeight: "300px",
    },
    dropdownItem: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const dacModalReadOnlyDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      marginTop: 5,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#000",
      border: "1px solid #000",
    },
    callout: {
      display: "none",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000", display: "none" },
  };
  const dacModalLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 240,
      fontSize: 13,
      color: "#323130",
    },
  };
  const dacIconStyleClass = mergeStyleSets({
    edit: {
      color: "#2392B2",
      fontSize: 20,
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      fontWeight: 600,
    },
    refresh: {
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

    popupCloseIcon: {
      color: "#ababab",
      fontSize: 18,
      fontWeight: 600,
      height: 20,
      width: 22,
      cursor: "pointer",
      backgroundColor: "#fff",
      padding: 5,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      ":hover": {
        color: "#f00",
      },
    },
  });
  // Styles-Section Ends
  // States-Declaration Starts
  const [dacReRender, setDacReRender] = useState<boolean>(true);
  const [dacUnsortMasterData, setDacUnsortMasterData] = useState<IData[]>([]);
  const [dacMasterData, setDacMasterData] = useState<IData[]>([]);
  const [dacData, setDacData] = useState<IData[]>([]);
  const [dacDisplayData, setDacDisplayData] = useState<IData[]>([]);
  const [daccurrentPage, setDacCurrentPage] = useState<number>(currentpage);
  const [dacDropDownOptions, setDacDropDownOptions] =
    useState<IDacDropDown>(dacDrpDwnOptns);
  const [dacFilters, setDacFilters] = useState<IDacFilterKeys>(dacFilterKeys);
  const [dacPopup, setDacPopup] = useState<{
    condition: boolean;
    responseData: IData;
    validation: boolean;
  }>({ condition: false, responseData: null, validation: false });
  const [dacLoader, setDacLoader] = useState<string>("noLoader");
  const [dacMasterColumns, setDacMasterColumns] =
    useState<IColumn[]>(dacColumns);
  // States-Declaration Ends
  //Function-Section Starts

  const getMasterUserData = (): void => {
    sharepointWeb.lists
      .getByTitle(MasterUserListName)
      .items.top(5000)
      .get()
      .then((muData) => {
        let positionDropDown: IDropdownOption[] = [];
        muData.forEach((_data) => {
          if (
            _data.Position &&
            positionDropDown.findIndex((_choi) => {
              return _choi.key == _data.Position;
            }) == -1
          ) {
            positionDropDown.push({
              key: _data.Position,
              text: _data.Position,
            });
          }
        });
        // positionDropDown.length > 0
        //   ? positionDropDown.unshift({ key: "Select", text: "Select" })
        //   : null;

        dacGetData(positionDropDown);
      })
      .catch((err) => {
        dacErrorFunction(err, "DAC-getMasterUserData");
      });
  };
  const dacGetData = (positionDropDown: IDropdownOption[]): void => {
    const loopingFunction = (peopleStr: string): IPeoplelist[] => {
      let peopleObj: IPeoplelist[] = [];

      if (peopleStr) {
        let peopleArr = peopleStr.split(",");
        if (peopleArr.length > 0) {
          peopleArr.forEach((people) => {
            let filteredValue: IPeoplelist[] = allPeoples.filter((obj) => {
              return obj.secondaryText == people;
            });
            filteredValue.length > 0 ? peopleObj.push(filteredValue[0]) : null;
          });
        }
      }

      return peopleObj;
    };
    sharepointWeb.lists
      .getByTitle(DLApprovalConfigListName)
      .items.select("*", "To/Id", "To/Title", "To/EMail")
      .expand("To")
      .orderBy("Modified", false)
      .top(5000)
      .get()
      .then((items) => {
        let dacAllitems: IData[] = [];
        // let dacAllitems = [];
        items.forEach((item) => {
          dacAllitems.push({
            ID: item.Id ? item.Id : "",
            ToId: item.ToId ? item.ToId : null,
            ToName: item.ToId ? item.To.Title : "",
            ToEmail: item.ToId ? item.To.EMail : "",

            BusinessArea: item.BusinessArea ? item.BusinessArea : "",

            // Choices
            EmployeesChoices: item.Employees
              ? item.Employees.split("~")[0]
              : "Not Mandatory",
            ContractorsChoices: item.Contractors
              ? item.Contractors.split("~")[0]
              : "Not Mandatory",
            ManagementChoices: item.Management
              ? item.Management.split("~")[0]
              : "Not Mandatory",
            BoardChoices: item.Board
              ? item.Board.split("~")[0]
              : "Not Mandatory",
            SchoolsChoices: item.Schools
              ? item.Schools.split("~")[0]
              : "Not Mandatory",
            PrincipalsChoices: item.Principals
              ? item.Principals.split("~")[0]
              : "Not Mandatory",
            InstructionCoachesChoices: item.InstructionCoaches
              ? item.InstructionCoaches.split("~")[0]
              : "Not Mandatory",
            TeachersChoices: item.Teachers
              ? item.Teachers.split("~")[0]
              : "Not Mandatory",
            TeachingAssistantsChoices: item.TeachingAssistants
              ? item.TeachingAssistants.split("~")[0]
              : "Not Mandatory",
            SchoolTeamsChoices: item.SchoolTeams
              ? item.SchoolTeams.split("~")[0]
              : "Not Mandatory",
            ParentsChoices: item.Parents
              ? item.Parents.split("~")[0]
              : "Not Mandatory",
            CommunitiesChoices: item.Communities
              ? item.Communities.split("~")[0]
              : "Not Mandatory",
            OtherChoices: item.Other
              ? item.Other.split("~")[0]
              : "Not Mandatory",
            ProfessionalsChoices: item.Professionals
              ? item.Professionals.split("~")[0]
              : "Not Mandatory",
            ServicesChoices: item.Services
              ? item.Services.split("~")[0]
              : "Not Mandatory",
            FundingChoices: item.Funding
              ? item.Funding.split("~")[0]
              : "Not Mandatory",
            ResourcesChoices: item.Resources
              ? item.Resources.split("~")[0]
              : "Not Mandatory",
            GovernmentsChoices: item.Governments
              ? item.Governments.split("~")[0]
              : "Not Mandatory",
            PoliticiansChoices: item.Politicians
              ? item.Politicians.split("~")[0]
              : "Not Mandatory",
            IndigenousChoices: item.Indigenous
              ? item.Indigenous.split("~")[0]
              : "Not Mandatory",
            ProfessionalAssociationsChoices: item.ProfessionalAssociations
              ? item.ProfessionalAssociations.split("~")[0]
              : "Not Mandatory",
            MediaChoices: item.Media
              ? item.Media.split("~")[0]
              : "Not Mandatory",

            // Position
            EmployeesPosition: item.Employees
              ? item.Employees.split("~")[1].split(",")
              : [],
            ContractorsPosition: item.Contractors
              ? item.Contractors.split("~")[1].split(",")
              : [],
            ManagementPosition: item.Management
              ? item.Management.split("~")[1].split(",")
              : [],
            BoardPosition: item.Board
              ? item.Board.split("~")[1].split(",")
              : [],
            SchoolsPosition: item.Schools
              ? item.Schools.split("~")[1].split(",")
              : [],
            PrincipalsPosition: item.Principals
              ? item.Principals.split("~")[1].split(",")
              : [],
            InstructionCoachesPosition: item.InstructionCoaches
              ? item.InstructionCoaches.split("~")[1].split(",")
              : [],
            TeachersPosition: item.Teachers
              ? item.Teachers.split("~")[1].split(",")
              : [],
            TeachingAssistantsPosition: item.TeachingAssistants
              ? item.TeachingAssistants.split("~")[1].split(",")
              : [],
            SchoolTeamsPosition: item.SchoolTeams
              ? item.SchoolTeams.split("~")[1].split(",")
              : [],
            ParentsPosition: item.Parents
              ? item.Parents.split("~")[1].split(",")
              : [],
            CommunitiesPosition: item.Communities
              ? item.Communities.split("~")[1].split(",")
              : [],
            OtherPosition: item.Other
              ? item.Other.split("~")[1].split(",")
              : [],
            ProfessionalsPosition: item.Professionals
              ? item.Professionals.split("~")[1].split(",")
              : [],
            ServicesPosition: item.Services
              ? item.Services.split("~")[1].split(",")
              : [],
            FundingPosition: item.Funding
              ? item.Funding.split("~")[1].split(",")
              : [],
            ResourcesPosition: item.Resources
              ? item.Resources.split("~")[1].split(",")
              : [],
            GovernmentsPosition: item.Governments
              ? item.Governments.split("~")[1].split(",")
              : [],
            PoliticiansPosition: item.Politicians
              ? item.Politicians.split("~")[1].split(",")
              : [],
            IndigenousPosition: item.Indigenous
              ? item.Indigenous.split("~")[1].split(",")
              : [],
            ProfessionalAssociationsPosition: item.ProfessionalAssociations
              ? item.ProfessionalAssociations.split("~")[1].split(",")
              : [],
            MediaPosition: item.Media
              ? item.Media.split("~")[1].split(",")
              : [],

            // People
            EmployeesPeople: item.Employees
              ? loopingFunction(item.Employees.split("~")[2])
              : [],
            ContractorsPeople: item.Contractors
              ? loopingFunction(item.Contractors.split("~")[2])
              : [],
            ManagementPeople: item.Management
              ? loopingFunction(item.Management.split("~")[2])
              : [],
            BoardPeople: item.Board
              ? loopingFunction(item.Board.split("~")[2])
              : [],
            SchoolsPeople: item.Schools
              ? loopingFunction(item.Schools.split("~")[2])
              : [],
            PrincipalsPeople: item.Principals
              ? loopingFunction(item.Principals.split("~")[2])
              : [],
            InstructionCoachesPeople: item.InstructionCoaches
              ? loopingFunction(item.InstructionCoaches.split("~")[2])
              : [],
            TeachersPeople: item.Teachers
              ? loopingFunction(item.Teachers.split("~")[2])
              : [],
            TeachingAssistantsPeople: item.TeachingAssistants
              ? loopingFunction(item.TeachingAssistants.split("~")[2])
              : [],
            SchoolTeamsPeople: item.SchoolTeams
              ? loopingFunction(item.SchoolTeams.split("~")[2])
              : [],
            ParentsPeople: item.Parents
              ? loopingFunction(item.Parents.split("~")[2])
              : [],
            CommunitiesPeople: item.Communities
              ? loopingFunction(item.Communities.split("~")[2])
              : [],
            OtherPeople: item.Other
              ? loopingFunction(item.Other.split("~")[2])
              : [],
            ProfessionalsPeople: item.Professionals
              ? loopingFunction(item.Professionals.split("~")[2])
              : [],
            ServicesPeople: item.Services
              ? loopingFunction(item.Services.split("~")[2])
              : [],
            FundingPeople: item.Funding
              ? loopingFunction(item.Funding.split("~")[2])
              : [],
            ResourcesPeople: item.Resources
              ? loopingFunction(item.Resources.split("~")[2])
              : [],
            GovernmentsPeople: item.Governments
              ? loopingFunction(item.Governments.split("~")[2])
              : [],
            PoliticiansPeople: item.Politicians
              ? loopingFunction(item.Politicians.split("~")[2])
              : [],
            IndigenousPeople: item.Indigenous
              ? loopingFunction(item.Indigenous.split("~")[2])
              : [],
            ProfessionalAssociationsPeople: item.ProfessionalAssociations
              ? loopingFunction(item.ProfessionalAssociations.split("~")[2])
              : [],
            MediaPeople: item.Media
              ? loopingFunction(item.Media.split("~")[2])
              : [],
          });
        });

        let sortedDropDownOptions: IDacDropDown = dacGetAllOptions(
          dacAllitems,
          positionDropDown
        );

        setDacDropDownOptions(sortedDropDownOptions);

        paginateFunction(1, dacAllitems);

        setDacUnsortMasterData([...dacAllitems]);
        columnSortArr = dacAllitems;
        setDacData([...dacAllitems]);
        columnSortMasterArr = dacAllitems;
        setDacMasterData([...dacAllitems]);
        setDacLoader("noLoader");
      })
      .catch((err) => {
        dacErrorFunction(err, "dacGetData");
      });
  };
  const dacGetAllOptions = (
    _data: IData[],
    positionDropDown: IDropdownOption[]
  ): IDacDropDown => {
    let _dacDropDown: IDacDropDown = dacDropDownOptions;
    _data.forEach((item: IData) => {
      if (
        item.BusinessArea &&
        _dacDropDown.businessAreaFilter.findIndex((_choi) => {
          return _choi.key == item.BusinessArea;
        }) == -1
      ) {
        _dacDropDown.businessAreaFilter.push({
          key: item.BusinessArea,
          text: item.BusinessArea,
        });
      }
    });

    _dacDropDown.allPosition = [...positionDropDown];

    let unsortedDropDownOptions: IDacDropDown =
      dacSortingDrpDwnOptns(_dacDropDown);

    return unsortedDropDownOptions;
  };
  const dacSortingDrpDwnOptns = (
    unsortedDropDownOptions: IDacDropDown
  ): IDacDropDown => {
    return unsortedDropDownOptions;
  };
  const dacListFilter = (key: string, option: any): void => {
    let arrBeforeFilter: IData[] = [...dacMasterData];
    let tempFilterKeys: IDacFilterKeys = { ...dacFilters };
    tempFilterKeys[key] = option;

    if (tempFilterKeys.businessArea != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.BusinessArea == tempFilterKeys.businessArea;
      });
    }
    if (tempFilterKeys.toUser) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.ToName.toLowerCase().includes(
          tempFilterKeys.toUser.toLowerCase()
        );
      });
    }

    paginateFunction(1, arrBeforeFilter);

    columnSortArr = arrBeforeFilter;
    setDacData([...columnSortArr]);
    setDacFilters({ ...tempFilterKeys });
  };
  const dacOnChangeHandler = (key: string, value: any): void => {
    let _dacPopup = { ...dacPopup };

    _dacPopup.validation = false;
    _dacPopup.responseData[key] = value;

    setDacPopup({ ..._dacPopup });
  };
  const validationFunction = (): void => {
    if (dacPopup.responseData.ToId) {
      dacUpdateData();
    } else {
      let _dacPopup = { ...dacPopup };
      _dacPopup.validation = true;
      setDacPopup({ ..._dacPopup });
      setDacLoader("noLoader");
    }
  };
  const dacUpdateData = (): void => {
    const jsonFunction = (key: string): string => {
      let finalStr: string = "";

      let emailArr: string[] = [];
      let emailStr: string = "";

      if (dacPopup.responseData[`${key}People`].length > 0) {
        dacPopup.responseData[`${key}People`].forEach((people) => {
          if (
            people.secondaryText &&
            emailArr.findIndex((email) => {
              return email == people.secondaryText;
            }) == -1
          ) {
            emailArr.push(people.secondaryText);
          }
        });
        emailStr = emailArr.join(",");
      }

      finalStr = `${
        dacPopup.responseData[`${key}Choices`]
          ? dacPopup.responseData[`${key}Choices`]
          : ""
      }~${
        dacPopup.responseData[`${key}Position`].length > 0
          ? dacPopup.responseData[`${key}Position`].join(",")
          : ""
      }~${emailStr}`;

      return finalStr;
    };
    let _responseData = {
      ToId: dacPopup.responseData.ToId,
      Employees: jsonFunction("Employees"),
      Contractors: jsonFunction("Contractors"),
      Management: jsonFunction("Management"),
      Board: jsonFunction("Board"),
      Schools: jsonFunction("Schools"),
      Principals: jsonFunction("Principals"),
      InstructionCoaches: jsonFunction("InstructionCoaches"),
      Teachers: jsonFunction("Teachers"),
      TeachingAssistants: jsonFunction("TeachingAssistants"),
      SchoolTeams: jsonFunction("SchoolTeams"),
      Parents: jsonFunction("Parents"),
      Communities: jsonFunction("Communities"),
      Other: jsonFunction("Other"),
      Professionals: jsonFunction("Professionals"),
      Services: jsonFunction("Services"),
      Funding: jsonFunction("Funding"),
      Resources: jsonFunction("Resources"),
      Governments: jsonFunction("Governments"),
      Politicians: jsonFunction("Politicians"),
      Indigenous: jsonFunction("Indigenous"),
      ProfessionalAssociations: jsonFunction("ProfessionalAssociations"),
      Media: jsonFunction("Media"),
    };

    sharepointWeb.lists
      .getByTitle(DLApprovalConfigListName)
      .items.getById(dacPopup.responseData.ID)
      .update(_responseData)
      .then(() => {
        setDacLoader("startUpLoader");
        setDacPopup({
          condition: false,
          responseData: null,
          validation: false,
        });
        setDacReRender(!dacReRender);
      })
      .catch((err) => {
        dacErrorFunction(err, "dacUpdateData");
      });
  };

  const paginateFunction = (pagenumber: number, data: IData[]): void => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems: IData[] = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setDacDisplayData(paginatedItems);
      setDacCurrentPage(pagenumber);
    } else {
      setDacDisplayData([]);
      setDacCurrentPage(1);
    }
  };
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempDacColumns = dacColumns;
    const newColumns: IColumn[] = tempDacColumns.slice();
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
    setDacData([...newDisplayData]);
    setDacMasterData([...newMasterData]);
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

  const dacErrorFunction = (error: string, functionName: string): void => {
    console.log(error, functionName);

    let response = {
      ComponentName: "Stock list",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setDacPopup({
          condition: false,
          responseData: null,
          validation: false,
        });
        setDacLoader("noLoader");
        ErrorPopup();
      }
    );
  };
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  const HTMLGenerator = (key: string): JSX.Element => {
    return (
      <div style={{ marginBottom: 10, marginRight: 10 }}>
        <div>
          <Label className={styles.dacPopupLabel}>{key}</Label>
        </div>
        <div
          style={{
            display: "flex",
            // flexWrap: "wrap",
            justifyContent: "space-between",
          }}
        >
          <div>
            <Label styles={dacModalLabelStyles}>Is approval mandatory ?</Label>
            <Toggle
              styles={{ root: { marginTop: 6 } }}
              checked={dacPopup.responseData[`${key}Choices`] == "Mandatory"}
              onText="Yes"
              offText="No"
              onChange={(ev) => {
                dacOnChangeHandler(
                  `${key}Choices`,
                  dacPopup.responseData[`${key}Choices`] == "Mandatory"
                    ? "Not Mandatory"
                    : "Mandatory"
                );
              }}
            />
          </div>
          <div>
            <Label styles={dacModalLabelStyles}>
              Select job position to be excluded
            </Label>
            <Dropdown
              multiSelect
              placeholder="Select an option"
              styles={dacModalDropdownStyles}
              options={dacDropDownOptions.allPosition}
              dropdownWidth={"auto"}
              selectedKeys={
                dacPopup.responseData[`${key}Position`].length > 0
                  ? [...dacPopup.responseData[`${key}Position`]]
                  : []
              }
              onChange={(
                e: React.FormEvent<HTMLDivElement>,
                option: any
              ): void => {
                if (option) {
                  let tempdacPopup = { ...dacPopup };
                  tempdacPopup.responseData[`${key}Position`] = option.selected
                    ? [
                        ...tempdacPopup.responseData[`${key}Position`],
                        option.key as string,
                      ]
                    : tempdacPopup.responseData[`${key}Position`].filter(
                        (key) => key !== option.key
                      );
                  tempdacPopup.responseData[`${key}Position`].sort();
                  setDacPopup({ ...tempdacPopup });
                }
              }}
            />
          </div>
          <div>
            <Label styles={dacModalLabelStyles}>
              Select people or group to be excluded
            </Label>
            <NormalPeoplePicker
              className={styles.dacPopupModalPeoplePicker}
              inputProps={{
                placeholder:
                  dacPopup.responseData[`${key}People`] &&
                  dacPopup.responseData[`${key}People`].length == 0
                    ? "Find People"
                    : "",
              }}
              styles={{
                root: {
                  selectors: {
                    selectors: {
                      ".ms-BasePicker-text": {
                        border: "1px solid #000",
                      },
                    },
                  },
                },
              }}
              onResolveSuggestions={GetUserDetails}
              selectedItems={dacPopup.responseData[`${key}People`]}
              onChange={(selectedUser) => {
                selectedUser.length != 0
                  ? dacOnChangeHandler(`${key}People`, selectedUser)
                  : dacOnChangeHandler(`${key}People`, []);
              }}
            />
          </div>
        </div>
      </div>
    );
  };
  //Function-Section Ends
  useEffect(() => {
    setDacLoader("startUpLoader");
    getMasterUserData();
  }, [dacReRender]);
  return (
    <>
      <div style={{ padding: "5px 10px" }}>
        {dacLoader == "startUpLoader" ? <CustomLoader /> : null}
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
          <div className={styles.dacHeaderSection} style={{ paddingBottom: 5 }}>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <div className={styles.dacHeader}>
                Distribution approval configuration
              </div>
            </div>

            {/* Filter-Section Starts */}
            <div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  // flexWrap: "wrap",
                }}
              >
                <div className={styles.ddSection}>
                  <div>
                    <Label styles={dacLabelStyles}>Business area</Label>
                    <Dropdown
                      placeholder="Select an option"
                      styles={
                        dacFilters.businessArea != "All"
                          ? dacActiveDropdownStyles
                          : dacDropdownStyles
                      }
                      options={dacDropDownOptions.businessAreaFilter}
                      dropdownWidth={"auto"}
                      selectedKey={dacFilters.businessArea}
                      onChange={(e, option: any): void => {
                        dacListFilter("businessArea", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label styles={dacLabelStyles}>Approver</Label>
                    <SearchBox
                      styles={
                        dacFilters.toUser
                          ? dacActiveSearchBoxStyles
                          : dacSearchBoxStyles
                      }
                      value={dacFilters.toUser}
                      onChange={(e, value): void => {
                        dacListFilter("toUser", value);
                      }}
                    />
                  </div>
                  <div>
                    <Icon
                      iconName="Refresh"
                      title="Click to reset"
                      className={dacIconStyleClass.refresh}
                      onClick={(): void => {
                        paginateFunction(1, dacUnsortMasterData);
                        columnSortArr = dacMasterData;
                        setDacData(dacMasterData);
                        columnSortMasterArr = dacMasterData;
                        setDacMasterData(dacMasterData);
                        setDacMasterColumns(dacColumns);
                        setDacFilters({ ...dacFilterKeys });
                      }}
                    />
                  </div>
                </div>
                <div>
                  <Label
                    style={{
                      marginTop: 25,
                      marginLeft: 10,
                      fontWeight: 500,
                      color: "#323130",
                      fontSize: 13,
                    }}
                  >
                    Number of records :{" "}
                    <span style={{ color: "#038387" }}>{dacData.length}</span>
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
              items={dacDisplayData}
              columns={dacMasterColumns}
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
          {dacData.length > 0 ? (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                margin: "10px 0",
              }}
            >
              <Pagination
                currentPage={daccurrentPage}
                totalPages={
                  dacData.length > 0
                    ? Math.ceil(dacData.length / totalPageItems)
                    : 1
                }
                onChange={(page): void => {
                  paginateFunction(page, dacData);
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
        {dacPopup.condition ? (
          <Modal
            isOpen={dacPopup.condition}
            styles={{
              root: {
                selectors: {
                  ".ms-Dialog-main": {
                    width: "60%",
                    overflow: "hidden",
                  },
                },
              },
            }}
            isBlocking={true}
          >
            <div style={{ padding: "15px 20px" }}>
              {/* Header-Section Starts */}
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  marginBottom: 10,
                }}
              >
                <Label
                  className={styles.dacHeaderPopupLabel}
                  style={{ marginLeft: "auto" }}
                >
                  Distribution approval
                </Label>
                <div style={{ marginLeft: "auto" }}>
                  <div>
                    <Icon
                      iconName="ChromeClose"
                      title="Close"
                      className={dacIconStyleClass.popupCloseIcon}
                      onClick={() => {
                        dacLoader == "noLoader"
                          ? setDacPopup({
                              condition: false,
                              responseData: null,
                              validation: false,
                            })
                          : null;
                      }}
                    />
                  </div>
                </div>
              </div>
              {/* Header-Section Ends */}
              {/* Body-Section Starts */}
              <div
                className={styles.dacPopupBodySection}
                // style={{ height: 500, overflow: "auto" }}
              >
                <div style={{ marginBottom: 10 }}>
                  <div>
                    <Label className={styles.dacPopupLabel}>General</Label>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      flexWrap: "wrap",
                      justifyContent: "space-between",
                    }}
                  >
                    <div>
                      <Label styles={dacModalLabelStyles}>Business area</Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={dacModalReadOnlyDropdownStyles}
                        options={[
                          {
                            key: dacPopup.responseData.BusinessArea,
                            text: dacPopup.responseData.BusinessArea,
                          },
                        ]}
                        dropdownWidth={"auto"}
                        selectedKey={dacPopup.responseData.BusinessArea}
                      />
                    </div>
                    <div>
                      <Label styles={dacModalLabelStyles}>Approver</Label>
                      <NormalPeoplePicker
                        className={styles.dacPopupModalPeoplePicker}
                        inputProps={{
                          placeholder: !dacPopup.responseData.ToId
                            ? "Find People"
                            : "",
                        }}
                        styles={
                          dacPopup.validation
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
                                    ".ms-BasePicker-text": {
                                      border: "1px solid #000",
                                    },
                                  },
                                },
                              }
                        }
                        onResolveSuggestions={GetUserDetails}
                        selectedItems={allPeoples.filter((people) => {
                          return (
                            people.ID ==
                            (dacPopup.responseData.ToId
                              ? dacPopup.responseData.ToId
                              : null)
                          );
                        })}
                        itemLimit={1}
                        onChange={(selectedUser) => {
                          selectedUser.length != 0
                            ? dacOnChangeHandler("ToId", selectedUser[0]["ID"])
                            : dacOnChangeHandler("ToId", null);
                        }}
                      />
                    </div>
                    <div style={{ width: 300 }}></div>
                  </div>
                </div>
                {HTMLGenerator("Employees")}
                {HTMLGenerator("Contractors")}
                {HTMLGenerator("Management")}
                {HTMLGenerator("Board")}
                {HTMLGenerator("Schools")}
                {HTMLGenerator("Principals")}
                {HTMLGenerator("InstructionCoaches")}
                {HTMLGenerator("Teachers")}
                {HTMLGenerator("TeachingAssistants")}
                {HTMLGenerator("SchoolTeams")}
                {HTMLGenerator("Parents")}
                {HTMLGenerator("Communities")}
                {HTMLGenerator("Other")}
                {HTMLGenerator("Professionals")}
                {HTMLGenerator("Services")}
                {HTMLGenerator("Funding")}
                {HTMLGenerator("Resources")}
                {HTMLGenerator("Governments")}
                {HTMLGenerator("Politicians")}
                {HTMLGenerator("Indigenous")}
                {HTMLGenerator("ProfessionalAssociations")}
                {HTMLGenerator("Media")}
              </div>
              {/* Body-Section Ends */}
              {/* Footer-Section Starts */}
              <div className={styles.dacPopupButtonSection}>
                {dacPopup.validation ? (
                  <Label style={{ color: "#f00" }}>
                    * Approver is mandatory
                  </Label>
                ) : null}
                <button
                  className={styles.PopupSuccessBtn}
                  onClick={() => {
                    if (dacLoader == "noLoader") {
                      setDacLoader("submitLoader");
                      validationFunction();
                    }
                  }}
                >
                  {dacLoader == "submitLoader" ? <Spinner /> : "Update"}
                </button>
                <button
                  className={styles.PopupCancelBtn}
                  onClick={() => {
                    dacLoader == "noLoader"
                      ? setDacPopup({
                          condition: false,
                          responseData: null,
                          validation: false,
                        })
                      : null;
                  }}
                >
                  Close
                </button>
              </div>
              {/* Footer-Section Ends */}
            </div>
          </Modal>
        ) : null}
        {/* Modal-Section Ends */}
      </div>
    </>
  );
};

export default DistributionApprovalConfig;
