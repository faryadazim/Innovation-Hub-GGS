import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import CustomLoader from "./CustomLoader";
import {
  Icon,
  Label,
  ILabelStyles,
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  Modal,
  IModalStyles,
  TextField,
  ITextFieldStyles,
  Checkbox,
  Spinner,
} from "@fluentui/react";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import { _Web } from "@pnp/sp/webs/types";

import Service from "../components/Services";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

interface IProps {
  context: WebPartContext;
  spcontext: any;
  graphContent: any;
  URL: string;
  handleclick: any;
  peopleList: any;
  isAdmin: boolean;
  docReviewDetails: IDocReviewDetails;
  distributionListHandlerFunction: any;
}
interface IDocReviewDetails {
  response: string;
  dlID: number;
  docReviewID: number;
  isEditable: boolean;
}
interface IAllPoliticiansData {
  state: string;
  senate: string;
  hor: string;
}
interface IData {
  isEmpty: boolean;
  ID: number;
  action: string;
  // GGSA
  innovationTeamChecked: boolean;
  innovationTeam: string[];
  employeesChecked: boolean;
  employees: string[];
  contractorsChecked: boolean;
  contractors: string[];
  BAChecked: boolean;
  BA: string[];
  managementTeamChecked: boolean;
  managementTeam: string[];
  boardChecked: boolean;
  board: string[];
  GGSAComments: string;

  // Schools
  schoolsSelectAll: boolean;

  schoolsChecked: boolean;
  schools: string[];
  principalsChecked: boolean;
  principals: string[];
  instructionCoachesChecked: boolean;
  instructionCoaches: string[];
  teachersChecked: boolean;
  teachers: string[];
  teachingAssistantsChecked: boolean;
  teachingAssistants: string[];
  schoolTeamChecked: boolean;
  schoolTeam: string[];
  parentsChecked: boolean;
  parents: string[];
  communityChecked: boolean;
  community: string[];
  hodChecked: boolean;
  hod: string[];
  schoolsOtherChecked: boolean;
  schoolsOther: string[];
  schoolsComments: string;

  // DigitalChannels
  gtpChecked: boolean;
  gtp: string[];
  digitalDatabaseChecked: boolean;
  digitalDatabase: string[];
  websiteChecked: boolean;
  website: string[];
  intranetChecked: boolean;
  intranet: string[];
  brochureStandChecked: boolean;
  brochureStand: string[];
  collateralCabinetChecked: boolean;
  collateralCabinet: string[];
  socialMediaChecked: boolean;
  socialMedia: string[];
  ceoLinkedInChecked: boolean;
  ceoLinkedIn: string[];
  productsChecked: boolean;
  products: string[];
  digitalChannelsOtherChecked: boolean;
  digitalChannelsOther: string[];
  digitalChannelsComments: string;

  // Partners
  partnersSelectAll: boolean;

  professionalsChecked: boolean;
  professionals: string[];
  servicesChecked: boolean;
  services: string[];
  fundingChecked: boolean;
  funding: string[];
  resourcesChecked: boolean;
  resources: string[];
  governmentsChecked: boolean;
  governments: string[];
  indigenousChecked: boolean;
  indigenous: string[];
  PAChecked: boolean;
  PA: string[];
  mediaChecked: boolean;
  media: string[];
  partnersComments: string;

  //Politicians
  politiciansSelectAll: boolean;
  politicians: IPoliticianData[];
}
interface IPoliticianData {
  SATChecked: boolean;
  SAT: string;

  senateChecked: boolean;
  senate: string[];

  HORChecked: boolean;
  HOR: string[];

  Comments: string;
}
interface IErrorStatus {
  // GGSA
  innovationTeam: boolean;
  employees: boolean;
  contractors: boolean;
  BA: boolean;
  managementTeam: boolean;
  board: boolean;

  GGSAOverAllValidation: boolean;

  // Schools
  schools: boolean;
  principals: boolean;
  instructionCoaches: boolean;
  teachers: boolean;
  teachingAssistants: boolean;
  schoolTeam: boolean;
  parents: boolean;
  community: boolean;
  hod: boolean;
  schoolsOther: boolean;

  SchoolsOverAllValidation: boolean;

  // DigitalChannels
  gtp: boolean;
  digitalDatabase: boolean;
  website: boolean;
  intranet: boolean;
  socialMedia: boolean;
  ceoLinkedIn: boolean;
  products: boolean;
  digitalChannelsOther: boolean;
  collateralCabinet: boolean;
  brochureStand: boolean;

  digitalChannelsOverAllValidation: boolean;

  // Partners
  professionals: boolean;
  services: boolean;
  funding: boolean;
  resources: boolean;
  governments: boolean;
  indigenous: boolean;
  PA: boolean;
  media: boolean;

  partnersOverAllValidation: boolean;

  //Politicians
  politicians: IPoliciansErrorStatus[];
  politiciansOverAllValidation: boolean;
}
interface IPoliciansErrorStatus {
  SAT: boolean;
  senate: boolean;
  HOR: boolean;
}
interface IDropDown {
  key: string | number;
  text: string;
}
interface IDropDownOptions {
  // GGSA
  innovationTeamOptns: IDropDown[];
  employeesOptns: IDropDown[];
  contractorsOptns: IDropDown[];
  BAOptns: IDropDown[];
  managementTeamOptns: IDropDown[];
  boardOptns: IDropDown[];

  // Schools
  schoolsOptns: IDropDown[];
  principalsOptns: IDropDown[];
  instructionCoachesOptns: IDropDown[];
  teachersOptns: IDropDown[];
  teachingAssistantsOptns: IDropDown[];
  schoolTeamOptns: IDropDown[];
  parentsOptns: IDropDown[];
  communityOptns: IDropDown[];
  hodOptns: IDropDown[];
  schoolsOtherOptns: IDropDown[];

  // Digital Channels
  gtpOptns: IDropDown[];
  digitalDatabaseOptns: IDropDown[];
  websiteOptns: IDropDown[];
  intranetOptns: IDropDown[];
  brochureStandOptns: IDropDown[];
  collateralCabinetOptns: IDropDown[];
  socialMediaOptns: IDropDown[];
  ceoLinkedInOptns: IDropDown[];
  productsOptns: IDropDown[];
  digitalChannelsOtherOptns: IDropDown[];

  // Partners
  professionalsOptns: IDropDown[];
  servicesOptns: IDropDown[];
  fundingOptns: IDropDown[];
  resourcesOptns: IDropDown[];
  governmentsOptns: IDropDown[];
  indigenousOptns: IDropDown[];
  PAOptns: IDropDown[];
  mediaOptns: IDropDown[];

  //Politicians
  SATOptns: IDropDown[];
  senateOptns: IDropDown[];
  HOROptns: IDropDown[];
}

let editableStatus: boolean;

const DistributionList = (props: IProps): JSX.Element => {
  const _webURL = Web(props.URL);
  const DRListName: string = "Review Log";
  const DLListName: string = "Distribution List";
  const DLAccessListName: string = "Distribution Access Level List";
  const DLApprovalConfigListName: string = "Distribution Approval Config";
  const MasterUserListName: string = "Master User List";

  let loggedUserEmail: string = props.spcontext.pageContext.user.email;
  let loggedUserID = props.peopleList.filter(
    (dev) => dev.secondaryText == loggedUserEmail
  )[0].ID;

  // variable-Declaration Starts
  const _responseData: IData = {
    isEmpty: true,
    ID: null,
    action: "new",
    // GGSA
    innovationTeamChecked: false,
    innovationTeam: [],
    employeesChecked: false,
    employees: [],
    contractorsChecked: false,
    contractors: [],
    BAChecked: false,
    BA: [],
    managementTeamChecked: false,
    managementTeam: [],
    boardChecked: false,
    board: [],
    GGSAComments: "",

    // Schools
    schoolsSelectAll: false,

    schoolsChecked: false,
    schools: [],
    principalsChecked: false,
    principals: [],
    instructionCoachesChecked: false,
    instructionCoaches: [],
    teachersChecked: false,
    teachers: [],
    teachingAssistantsChecked: false,
    teachingAssistants: [],
    schoolTeamChecked: false,
    schoolTeam: [],
    parentsChecked: false,
    parents: [],
    communityChecked: false,
    community: [],
    hodChecked: false,
    hod: [],
    schoolsOtherChecked: false,
    schoolsOther: [],
    schoolsComments: "",

    // Digital Channels
    gtpChecked: false,
    gtp: [],
    digitalDatabaseChecked: false,
    digitalDatabase: [],
    websiteChecked: false,
    website: [],
    intranetChecked: false,
    intranet: [],
    brochureStandChecked: false,
    brochureStand: [],
    collateralCabinetChecked: false,
    collateralCabinet: [],
    socialMediaChecked: false,
    socialMedia: [],
    ceoLinkedInChecked: false,
    ceoLinkedIn: [],
    productsChecked: false,
    products: [],
    digitalChannelsOtherChecked: false,
    digitalChannelsOther: [],
    digitalChannelsComments: "",

    // Partners
    partnersSelectAll: false,

    professionalsChecked: false,
    professionals: [],
    servicesChecked: false,
    services: [],
    fundingChecked: false,
    funding: [],
    resourcesChecked: false,
    resources: [],
    governmentsChecked: false,
    governments: [],
    indigenousChecked: false,
    indigenous: [],
    PAChecked: false,
    PA: [],
    mediaChecked: false,
    media: [],
    partnersComments: "",

    //Politicians
    politiciansSelectAll: false,
    politicians: [
      {
        SATChecked: false,
        SAT: "",

        senateChecked: false,
        senate: [],

        HORChecked: false,
        HOR: [],

        Comments: "",
      },
    ],
  };
  const _responseErrorStatus: IErrorStatus = {
    // GGSA
    innovationTeam: false,
    employees: false,
    contractors: false,
    BA: false,
    managementTeam: false,
    board: false,

    GGSAOverAllValidation: false,

    // Schools
    schools: false,
    principals: false,
    instructionCoaches: false,
    teachers: false,
    teachingAssistants: false,
    schoolTeam: false,
    parents: false,
    community: false,
    hod: false,
    schoolsOther: false,

    SchoolsOverAllValidation: false,

    // Digital Channels
    gtp: false,
    digitalDatabase: false,
    website: false,
    intranet: false,
    brochureStand: false,
    collateralCabinet: false,
    socialMedia: false,
    ceoLinkedIn: false,
    products: false,
    digitalChannelsOther: false,

    digitalChannelsOverAllValidation: false,

    // Partners
    professionals: false,
    services: false,
    funding: false,
    resources: false,
    governments: false,
    indigenous: false,
    PA: false,
    media: false,

    partnersOverAllValidation: false,

    //Politicians
    politicians: [
      {
        SAT: false,
        senate: false,
        HOR: false,
      },
    ],
    politiciansOverAllValidation: false,
  };
  const _DrpDwnOptns: IDropDownOptions = {
    // GGSA
    innovationTeamOptns: [],
    employeesOptns: [],
    contractorsOptns: [],
    BAOptns: [],
    managementTeamOptns: [],
    boardOptns: [],

    // Schools
    schoolsOptns: [],
    principalsOptns: [],
    instructionCoachesOptns: [],
    teachersOptns: [],
    teachingAssistantsOptns: [],
    schoolTeamOptns: [],
    parentsOptns: [],
    communityOptns: [],
    hodOptns: [],
    schoolsOtherOptns: [],

    // Digital Channels
    gtpOptns: [],
    digitalDatabaseOptns: [],
    websiteOptns: [],
    intranetOptns: [],
    brochureStandOptns: [],
    collateralCabinetOptns: [],
    socialMediaOptns: [],
    ceoLinkedInOptns: [],
    productsOptns: [],
    digitalChannelsOtherOptns: [],

    // Partners
    professionalsOptns: [],
    servicesOptns: [],
    fundingOptns: [],
    resourcesOptns: [],
    governmentsOptns: [],
    indigenousOptns: [],
    PAOptns: [],
    mediaOptns: [],

    //Politicians
    SATOptns: [],
    senateOptns: [],
    HOROptns: [],
  };
  // variable-Declaration Ends

  // Style-Section Starts
  const headingStyles: Partial<ILabelStyles> = {
    root: {
      color: "#000",
      fontSize: 24,
      padding: 0,
      marginBottom: 10,
    },
  };
  const generalTitleStyles: Partial<ILabelStyles> = {
    root: {
      color: "#000",
      fontSize: 16,
    },
  };
  const DLDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 400,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      color: "#000",
      border: "1px solid #E8E8EA",
      borderRadius: 4,
    },
    dropdownItem: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    callout: {
      minHeight: "100px !important",
      maxHeight: "300px !important",
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const DLErrorDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 400,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      color: "#000",
      border: "2px solid #f00",
      borderRadius: 4,
    },
    dropdownItem: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    callout: {
      minHeight: "100px !important",
      maxHeight: "300px !important",
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const DLMultiTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 400,
      borderRadius: 4,
    },
    field: { fontSize: 14, color: "#000" },
    fieldGroup: { border: "1px solid #E8E8EA" },
  };
  const DLModalStyles: Partial<IModalStyles> = {
    root: { borderRadius: "none" },
    main: {
      width: 500,
      margin: 10,
      padding: "20px 10px",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
    },
  };
  const DLGeneralStyles = mergeStyleSets({
    headingWithCheckbox: {
      display: "flex",
      width: 400,
      marginBottom: 10,
      marginRight: 10,
    },
  });
  const DLIconStyleClass = mergeStyleSets({
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
    backIcon: {
      cursor: "pointer",
      color: "#2392b2",
      fontSize: 24,
      marginTop: 6,
      marginRight: 12,
    },
    circleAddIcon: {
      fontSize: 17,
      height: 14,
      width: 17,
      cursor: "pointer",
      color: "#2392b2",
      marginTop: 0,
      marginRight: 12,
    },
    deleteIcon: {
      fontSize: 17,
      height: 14,
      width: 17,
      color: "#fff",
      marginTop: 0,
      marginRight: 5,
      marginLeft: 5,
      fontWeight: 600,
      userSelect: "none",
    },
  });
  // Style-Section Ends

  // State-Declaration Starts
  const [DLReRender, setDLReRender] = useState<boolean>(false);
  const [DLApproverData, setDLApproverData] = useState<{
    userData: any;
    dlConfigData: any;
  }>(null);
  const [DLUserPermission, setDLUserPermission] = useState<string>("");
  const [DLPageSwitch, setDLPageSwitch] = useState("GGSA");
  const [DLPoliticiansData, setDLPoliticiansData] = useState<
    IAllPoliticiansData[]
  >([]);
  const [DLResponseData, setDLResponseData] = useState<IData>(_responseData);
  const [DLResponseErrorStatus, setDLResponseErrorStatus] =
    useState<IErrorStatus>(_responseErrorStatus);
  const [DLDropdownOptions, setDLDropdownOptions] =
    useState<IDropDownOptions>(_DrpDwnOptns);
  const [DisableApproveBtn, setDisableApproveBtn] = useState<boolean>(false);
  const [DLWarningPopup, setDLWarningPopup] = useState<boolean>(false);
  const [mailTemplatePopup, setMailTemplatePopup] = useState<boolean>(false);
  const [DLLoader, setDLLoader] = useState<string>("noLoader");
  // State-Declaration Ends

  // Function-Declaration Starts
  const HTMLBulider = (
    section: string,
    labelText: string,
    key: string
  ): JSX.Element => {
    let index: string;
    let _key: string;
    if (section == "Politicians") {
      index = key.split("-")[0];
      _key = key.split("-")[1];
    }
    return section == "Politicians" ? (
      _key == "SAT" ? (
        <div style={{ marginBottom: 15 }}>
          <div className={DLGeneralStyles.headingWithCheckbox}>
            <Checkbox
              styles={{ root: { marginTop: 6 } }}
              disabled={editableStatus == true ? false : true}
              checked={DLResponseData.politicians[index][`${_key}Checked`]}
              onChange={(ev) => {
                DL_onChangeHandler(
                  section,
                  "checkbox",
                  `${index}-${_key}Checked-${_key}`,
                  !DLResponseData.politicians[index][`${_key}Checked`]
                );
              }}
            />
            <Label styles={generalTitleStyles}>{labelText}</Label>
          </div>
          <Dropdown
            placeholder="Select an option"
            multiSelect={false}
            disabled={
              editableStatus == true &&
                DLResponseData.politicians[index][`${_key}Checked`] == true
                ? false
                : true
            }
            // disabled={true}
            styles={
              DLResponseErrorStatus.politicians[index][_key]
                ? DLErrorDropdownStyles
                : DLDropdownStyles
            }
            options={DLDropdownOptions[`${_key}Optns`]}
            selectedKey={DLResponseData.politicians[index][_key]}
            onChange={(
              e: React.FormEvent<HTMLDivElement>,
              item: IDropdownOption
            ) => {
              DL_onChangeHandler(
                section,
                "Dropdown",
                `${index}-${_key}`,
                item.key
              );
            }}
          />
        </div>
      ) : (
        <div style={{ marginBottom: 15 }}>
          <div className={DLGeneralStyles.headingWithCheckbox}>
            <Checkbox
              styles={{ root: { marginTop: 6 } }}
              disabled={editableStatus == true ? false : true}
              checked={DLResponseData.politicians[index][`${_key}Checked`]}
              onChange={(ev) => {
                DL_onChangeHandler(
                  section,
                  "checkbox",
                  `${index}-${_key}Checked-${_key}`,
                  !DLResponseData.politicians[index][`${_key}Checked`]
                );
              }}
            />
            <Label styles={generalTitleStyles}>{labelText}</Label>
          </div>
          <Dropdown
            placeholder="Select an option"
            multiSelect
            disabled={
              editableStatus == true &&
                DLResponseData.politicians[index][`${_key}Checked`] == true
                ? false
                : true
            }
            // disabled={true}
            styles={
              DLResponseErrorStatus.politicians[index][_key]
                ? DLErrorDropdownStyles
                : DLDropdownStyles
            }
            // options={DLDropdownOptions[`${_key}Optns`]}
            options={DL_politicianDropDown(
              `${DLResponseData.politicians[index].SAT}`,
              _key
            )}
            selectedKeys={
              DLResponseData.politicians[index][_key].length > 0
                ? [...DLResponseData.politicians[index][_key]]
                : []
            }
            onChange={(
              e: React.FormEvent<HTMLDivElement>,
              item: IDropdownOption
            ) => {
              DL_onChangeHandler(section, "Dropdown", `${index}-${_key}`, item);
            }}
          />
        </div>
      )
    ) : section == "DigitalChannels" &&
      key != "products" &&
      key != "digitalChannelsOther" ? (
      <div style={{ marginBottom: 15 }}>
        <div className={DLGeneralStyles.headingWithCheckbox}>
          <Checkbox
            disabled={editableStatus == true ? false : true}
            styles={{ root: { marginTop: 6 } }}
            checked={DLResponseData[`${key}Checked`]}
            onChange={(ev) => {
              DL_onChangeHandler(
                section,
                "checkbox",
                `${key}Checked-${key}`,
                !DLResponseData[`${key}Checked`]
              );
            }}
          />
          <Label styles={generalTitleStyles}>{labelText}</Label>
        </div>
        <Dropdown
          placeholder="Select an option"
          // multiSelect
          disabled={
            editableStatus == true && DLResponseData[`${key}Checked`] == true
              ? false
              : true
          }
          // disabled={true}
          styles={
            DLResponseErrorStatus[key]
              ? DLErrorDropdownStyles
              : DLDropdownStyles
          }
          options={DLDropdownOptions[`${key}Optns`]}
          selectedKey={
            DLResponseData[key].length > 0 ? DLResponseData[key][0] : ""
          }
          onChange={(
            e: React.FormEvent<HTMLDivElement>,
            item: IDropdownOption
          ) => {
            DL_onChangeHandler(section, "singleDropdown", key, item);
          }}
        />
      </div>
    ) : (
      <div style={{ marginBottom: 15 }}>
        <div className={DLGeneralStyles.headingWithCheckbox}>
          <Checkbox
            disabled={editableStatus == true ? false : true}
            styles={{ root: { marginTop: 6 } }}
            checked={DLResponseData[`${key}Checked`]}
            onChange={(ev) => {
              DL_onChangeHandler(
                section,
                "checkbox",
                `${key}Checked-${key}`,
                !DLResponseData[`${key}Checked`]
              );
            }}
          />
          <Label styles={generalTitleStyles}>{labelText}</Label>
        </div>
        <Dropdown
          placeholder="Select an option"
          multiSelect
          disabled={
            editableStatus == true && DLResponseData[`${key}Checked`] == true
              ? false
              : true
          }
          // disabled={true}
          styles={
            DLResponseErrorStatus[key]
              ? DLErrorDropdownStyles
              : DLDropdownStyles
          }
          options={DLDropdownOptions[`${key}Optns`]}
          selectedKeys={
            DLResponseData[key].length > 0 ? [...DLResponseData[key]] : []
          }
          onChange={(
            e: React.FormEvent<HTMLDivElement>,
            item: IDropdownOption
          ) => {
            DL_onChangeHandler(section, "Dropdown", key, item);
          }}
        />
      </div>
    );
  };
  const textFieldHTMLBuldier = (
    section: string,
    labelText: string,
    key: string
  ): JSX.Element => {
    let index: string;
    let _key: string;
    if (section == "Politicians") {
      index = key.split("-")[0];
      _key = key.split("-")[1];
    }
    return section == "Politicians" ? (
      <div style={{ marginBottom: 15 }}>
        <div className={DLGeneralStyles.headingWithCheckbox}>
          <Label styles={generalTitleStyles}>{labelText}</Label>
        </div>
        <TextField
          placeholder="Add Comments"
          readOnly={editableStatus == true ? false : true}
          value={DLResponseData.politicians[index][_key]}
          multiline
          rows={5}
          resizable={false}
          styles={DLMultiTxtBoxStyles}
          onChange={(e, value: string) => {
            DL_onChangeHandler(section, "text", `${index}-${_key}`, value);
          }}
        />
      </div>
    ) : (
      <div style={{ marginBottom: 15 }}>
        <div className={DLGeneralStyles.headingWithCheckbox}>
          <Label styles={generalTitleStyles}>{labelText}</Label>
        </div>
        <TextField
          placeholder="Add Comments"
          readOnly={editableStatus == true ? false : true}
          value={DLResponseData[key]}
          multiline
          rows={5}
          resizable={false}
          styles={DLMultiTxtBoxStyles}
          onChange={(e, value: string) => {
            DL_onChangeHandler(section, "text", key, value);
          }}
        />
      </div>
    );
  };
  const selectAllHTMLBuldier = (
    section: string,
    labelText: string,
    key: string
  ): JSX.Element => {
    return (
      <div className={styles.dlSelectAllSection}>
        <Checkbox
          styles={{
            root: { marginTop: 6 },
          }}
          disabled={editableStatus == true ? false : true}
          checked={DLResponseData[key]}
          onChange={(ev) => {
            DL_onChangeHandler(section, "selectAll", key, !DLResponseData[key]);
          }}
        />
        <Label styles={generalTitleStyles}>{labelText}</Label>
      </div>
    );
  };

  const GGSA_Tab = (): JSX.Element => {
    return (
      <>
        {DL_Footer(
          "GGSAOverAllValidation",
          "GGSA",
          DLUserPermission == "MediumPerms" ? "DigitalChannels" : "Schools"
        )}
        <div className={styles.flexWrapper}>
          {DLUserPermission == "FullPerms" ||
            DLUserPermission == "HighPerms" ||
            DLUserPermission == "MediumPerms"
            ? HTMLBulider("GGSA", "Innovation team", "innovationTeam")
            : null}
          {DLUserPermission == "FullPerms" || DLUserPermission == "HighPerms" ||
            DLUserPermission == "MediumPerms"
            ? HTMLBulider("GGSA", "Employees", "employees")
            : null}
          {DLUserPermission == "FullPerms" ||
            DLUserPermission == "HighPerms" ||
            DLUserPermission == "MediumPerms"
            ? HTMLBulider("GGSA", "Contractors", "contractors")
            : null}
          {DLUserPermission == "FullPerms" || DLUserPermission == "HighPerms" ||
            DLUserPermission == "MediumPerms"
            ? HTMLBulider("GGSA", "Business area", "BA")
            : null}
          {DLUserPermission == "FullPerms" || DLUserPermission == "HighPerms" ||
            DLUserPermission == "MediumPerms"
            ? HTMLBulider("GGSA", "Management team", "managementTeam")
            : null}
          {DLUserPermission == "FullPerms"
            ? HTMLBulider("GGSA", "Board", "board")
            : null}
          {DLUserPermission == "FullPerms" || DLUserPermission == "HighPerms"
            ||
            DLUserPermission == "MediumPerms"
            ? textFieldHTMLBuldier("GGSA", "Comments", "GGSAComments")
            : null}
        </div>
      </>
    );
  };
  const Schools_Tab = (): JSX.Element => {
    return (
      <>
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          {selectAllHTMLBuldier("Schools", "Select all", "schoolsSelectAll")}
          {DL_Footer(
            "SchoolsOverAllValidation",
            "Schools",
            DLUserPermission == "MediumPerms" ? "End" : "DigitalChannels"
          )}
        </div>
        <div className={styles.flexWrapper}>
          {HTMLBulider("Schools", "Principals", "principals")}
          {HTMLBulider("Schools", "Instruction coaches", "instructionCoaches")}
          {HTMLBulider("Schools", "Teachers", "teachers")}
          {HTMLBulider("Schools", "Teaching assistants", "teachingAssistants")}
          {HTMLBulider("Schools", "Schools", "schools")}
          {HTMLBulider("Schools", "School team", "schoolTeam")}
          {HTMLBulider("Schools", "Parents", "parents")}
          {HTMLBulider("Schools", "Communities", "community")}
          {HTMLBulider("Schools", "Head of Departments", "hod")}
          {HTMLBulider("Schools", "Other", "schoolsOther")}
          {textFieldHTMLBuldier("Schools", "Comments", "schoolsComments")}
        </div>
      </>
    );
  };
  const DigitalChannels_Tab = (): JSX.Element => {
    return (
      <>

        {/* {DL_Footer(
          "GGSAOverAllValidation",
          "GGSA",
          DLUserPermission == "MediumPerms" ? "End" : "Schools"
        )} */}




        {DL_Footer(
          "digitalChannelsOverAllValidation",
          "DigitalChannels",
          DLUserPermission == "MediumPerms" ? "End" : "Partners"

        )}
        <div className={styles.flexWrapper}>
          {HTMLBulider("DigitalChannels", "Great teaching portal", "gtp")}
          {HTMLBulider(
            "DigitalChannels",
            "Digital database",
            "digitalDatabase"
          )}
          {HTMLBulider("DigitalChannels", "Website", "website")}
          {HTMLBulider("DigitalChannels", "Intranet", "intranet")}
          {HTMLBulider(
            "DigitalChannels",
            "Social media (Facebook ,LinkedIn, Twitter)",
            "socialMedia"
          )}
          {HTMLBulider("DigitalChannels", "CEO LinkedIn", "ceoLinkedIn")}
          {HTMLBulider("DigitalChannels", "Products", "products")}
          {HTMLBulider("DigitalChannels", "Other", "digitalChannelsOther")}
          {HTMLBulider(
            "DigitalChannels",
            "Collateral cabinet",
            "collateralCabinet"
          )}
          {HTMLBulider("DigitalChannels", "Brochure stand", "brochureStand")}
          {textFieldHTMLBuldier(
            "DigitalChannels",
            "Comments",
            "digitalChannelsComments"
          )}
        </div>
      </>
    );
  };
  const Partners_Tab = (): JSX.Element => {
    return (
      <>
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          {selectAllHTMLBuldier("Partners", "Select all", "partnersSelectAll")}
          {DL_Footer(
            "partnersOverAllValidation",
            "Partners",
            DLUserPermission == "HighPerms" ? "End" : "Politicians"
          )}
        </div>
        <div className={styles.flexWrapper}>
          {HTMLBulider("Partners", "Professionals", "professionals")}
          {HTMLBulider("Partners", "Services", "services")}
          {DLUserPermission != "HighPerms"
            ? HTMLBulider("Partners", "Funding", "funding")
            : null}
          {HTMLBulider("Partners", "Resources", "resources")}
          {HTMLBulider("Partners", "Governments", "governments")}
          {HTMLBulider("Partners", "Indigenous", "indigenous")}
          {HTMLBulider("Partners", "Professional associations", "PA")}
          {HTMLBulider("Partners", "Media", "media")}
          {textFieldHTMLBuldier("Partners", "Comments", "partnersComments")}
        </div>
      </>
    );
  };
  const Politicians_Tab = (): JSX.Element => {
    return (
      <>
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          <div className={styles.dlPoliticianAddBtnSection}>
            <button
              className={
                editableStatus == true
                  ? styles.activeButton
                  : styles.inActiveButton
              }
              onClick={() => {
                if (editableStatus == true) {
                  let tempDLResponseData = { ...DLResponseData };
                  let tempDLResponseErrorStatus = {
                    ...DLResponseErrorStatus,
                  };
                  tempDLResponseData.politicians.push({
                    SATChecked: false,
                    SAT: "",
                    senateChecked: false,
                    senate: [],
                    HORChecked: false,
                    HOR: [],
                    Comments: "",
                  });
                  tempDLResponseErrorStatus.politicians.push({
                    SAT: false,
                    senate: false,
                    HOR: false,
                  });
                  setDLResponseData({ ...tempDLResponseData });
                  setDLResponseErrorStatus({
                    ...tempDLResponseErrorStatus,
                  });
                }
              }}
            >
              Add
            </button>
          </div>
          {DL_Footer("politiciansOverAllValidation", "Politicians", "End")}
        </div>
        <div style={{ height: 400, overflow: "auto", marginTop: 20 }}>
          {DLResponseData.politicians.map(
            (data: IPoliticianData, index: number) => {
              return (
                <div className={styles.flexWrapper}>
                  <div
                    className={
                      editableStatus == true
                        ? DLResponseData.politicians.length == 1
                          ? styles.dlPoliticianDeleteBtnSectionInActive
                          : styles.dlPoliticianDeleteBtnSectionActive
                        : styles.dlPoliticianDeleteBtnSectionInActive
                    }
                    onClick={() => {
                      if (
                        editableStatus == true &&
                        DLResponseData.politicians.length > 1
                      ) {
                        let tempDLResponseData = { ...DLResponseData };
                        let tempDLResponseErrorStatus = {
                          ...DLResponseErrorStatus,
                        };

                        tempDLResponseData.politicians.splice(index, 1);
                        tempDLResponseErrorStatus.politicians.slice(index, 1);
                        setDLResponseData({ ...tempDLResponseData });
                        setDLResponseErrorStatus({
                          ...tempDLResponseErrorStatus,
                        });
                      }
                    }}
                  >
                    <Icon
                      iconName="Delete"
                      className={DLIconStyleClass.deleteIcon}
                    />
                  </div>
                  {HTMLBulider(
                    "Politicians",
                    "State and territory",
                    `${index}-SAT`
                  )}
                  {data.SATChecked
                    ? HTMLBulider("Politicians", "Senate", `${index}-senate`)
                    : null}
                  {data.SATChecked
                    ? HTMLBulider(
                      "Politicians",
                      "House of representatives",
                      `${index}-HOR`
                    )
                    : null}
                  {data.SATChecked
                    ? textFieldHTMLBuldier(
                      "Politicians",
                      "Comments",
                      `${index}-Comments`
                    )
                    : null}
                </div>
              );
            }
          )}
        </div>
      </>
    );
  };

  const DL_Footer = (
    key: string,
    fromTab: string,
    toTab: string
  ): JSX.Element => {
    return (
      <div className={styles.dlFooterSection}>
        {toTab == "End" ? (
          editableStatus == true ? (
            <div style={{ display: "flex" }}>
              {key && DLResponseErrorStatus[key] ? (
                <Label
                  style={{ color: "#f00" }}
                >{`Please fill mandatory fields`}</Label>
              ) : null}
              {props.docReviewDetails.response == "Signed Off" ? (
                <button
                  className={
                    DisableApproveBtn ? styles.inActiveBtn : styles.activeBtn
                  }
                  onClick={() => {
                    if (DisableApproveBtn == false && DLLoader == "noLoader") {
                      setDLLoader("sendForApprovalLoader");
                      DL_validationFunction(
                        fromTab,
                        toTab == "End" ? "SendForApproval" : toTab
                      );
                    }
                  }}
                >
                  {DLLoader == "sendForApprovalLoader" ? (
                    <Spinner />
                  ) : (
                    "Send for approval"
                  )}
                </button>
              ) : (
                <button
                  className={
                    DisableApproveBtn ? styles.inActiveBtn : styles.activeBtn
                  }
                  onClick={() => {
                    if (DisableApproveBtn == false && DLLoader == "noLoader") {
                      setDLLoader("Submit");
                      DL_validationFunction(
                        fromTab,
                        toTab == "End" ? "PRSubmit" : toTab
                      );
                    }
                  }}
                >
                  {DLLoader == "Submit" ? <Spinner /> : "Submit"}
                </button>
              )}
              <button
                className={styles.activeBtn}
                onClick={() => {
                  if (DLLoader == "noLoader") {
                    setDLLoader("submitLoader");
                    DL_validationFunction(
                      fromTab,
                      toTab == "End" ? "Submit" : toTab
                    );
                  }
                }}
              >
                {DLLoader == "submitLoader" ? (
                  <Spinner />
                ) : DLResponseData.action == "new" ? (
                  "Save"
                ) : (
                  "Update"
                )}
              </button>
            </div>
          ) : null
        ) : (
          <div style={{ display: "flex", height: 30 }}>
            {key && DLResponseErrorStatus[key] ? (
              <Label
                style={{ color: "#f00" }}
              >{`Please fill mandatory fields`}</Label>
            ) : null}

            <button
              title={`Go to ${toTab}`}
              className={styles.activeBtn}
              onClick={() => {
                DL_validationFunction(fromTab, toTab);
              }}
            >
              Next
            </button>
          </div>
        )}
      </div>
    );
  };

  const DL_onChangeHandler = (
    section: string,
    type: string,
    key: string,
    value: any
  ): void => {
    let tempResponseData: IData = { ...DLResponseData };
    let tempResponseErrorStatus: IErrorStatus = { ...DLResponseErrorStatus };
    let tempDLDropdownOptions: IDropDownOptions = { ...DLDropdownOptions };

    tempResponseErrorStatus[
      section == "GGSA"
        ? "GGSAOverAllValidation"
        : section == "Schools"
          ? "SchoolsOverAllValidation"
          : section == "DigitalChannels"
            ? "digitalChannelsOverAllValidation"
            : section == "Partners"
              ? "partnersOverAllValidation"
              : section == "Politicians"
                ? "politiciansOverAllValidation"
                : null
    ] = false;

    if (type == "Dropdown") {
      if (section == "Politicians") {
        let index: string = key.split("-")[0];
        let _key: string = key.split("-")[1];
        if (_key == "SAT") {
          tempDLDropdownOptions.senateOptns = _DrpDwnOptns.senateOptns;
          tempDLDropdownOptions.HOROptns = _DrpDwnOptns.HOROptns;

          let fitleredArr = DLPoliticiansData.filter((_item) => {
            return _item.state == value;
          });

          fitleredArr.forEach((arr) => {
            if (arr.senate) {
              if (
                tempDLDropdownOptions.senateOptns.findIndex((optn) => {
                  return optn.key == arr.senate;
                }) == -1
              ) {
                tempDLDropdownOptions.senateOptns.push({
                  key: arr.senate,
                  text: arr.senate,
                });
              }
            }

            if (arr.hor) {
              if (
                tempDLDropdownOptions.HOROptns.findIndex((optn) => {
                  return optn.key == arr.hor;
                }) == -1
              ) {
                tempDLDropdownOptions.HOROptns.push({
                  key: arr.hor,
                  text: arr.hor,
                });
              }
            }
          });

          tempResponseData.politicians[index].SAT = value;
          tempResponseErrorStatus.politicians[index].SAT = false;

          tempResponseData.politicians[index].senate = [];
          tempResponseData.politicians[index].senateChecked = false;

          tempResponseData.politicians[index].HOR = [];
          tempResponseData.politicians[index].HORChecked = false;

          tempResponseErrorStatus.politicians[index].senate = false;
          tempResponseErrorStatus.politicians[index].HOR = false;
        } else {
          if (value) {
            if (value.key === "All" && value.selected) {
              tempResponseData.politicians[index][_key] = DLDropdownOptions[
                `${_key}Optns`
              ].map((option) => option.key as string);
              if (tempResponseData.politicians[index][_key][0] == "All") {
                tempResponseData.politicians[index][_key].shift();
                tempResponseData.politicians[index][_key].push("All");
              }
            } else if (value.key === "All") {
              tempResponseData.politicians[index][_key] = [];
            } else if (value.selected) {
              const newKeys = [value.key as string];
              if (
                tempResponseData.politicians[index][_key].length ===
                DLDropdownOptions[`${_key}Optns`].length - 2
              ) {
                newKeys.push("All");
              }
              tempResponseData.politicians[index][_key] = [
                ...tempResponseData.politicians[index][_key],
                ...newKeys,
              ];
            } else {
              tempResponseData.politicians[index][_key] =
                tempResponseData.politicians[index][_key].filter(
                  (key) => key !== value.key && key !== "All"
                );
            }
          }
          tempResponseErrorStatus.politicians[index][_key] = false;
        }
      } else {
        if (value) {
          if (value.key === "All" && value.selected) {
            tempResponseData[key] = DLDropdownOptions[`${key}Optns`].map(
              (option) => option.key as string
            );
            if (tempResponseData[key][0] == "All") {
              tempResponseData[key].shift();
              tempResponseData[key].push("All");
            }
          } else if (value.key === "All") {
            tempResponseData[key] = [];
          } else if (value.selected) {
            const newKeys = [value.key as string];
            if (
              tempResponseData[key].length ===
              DLDropdownOptions[`${key}Optns`].length - 2
            ) {
              newKeys.push("All");
            }
            tempResponseData[key] = [...tempResponseData[key], ...newKeys];
          } else {
            tempResponseData[key] = tempResponseData[key].filter(
              (key) => key !== value.key && key !== "All"
            );
          }
        }
        tempResponseErrorStatus[key] = false;
      }
    } else if (type == "checkbox") {
      if (section == "Politicians") {
        let index: string = key.split("-")[0];
        let _key: string = key.split("-")[1];
        let resetDataOf: string = key.split("-")[2];

        if (_key == "SATChecked") {
          tempResponseData.politicians[index].senateChecked = false;
          tempResponseData.politicians[index].senate = [];

          tempResponseData.politicians[index].HORChecked = false;
          tempResponseData.politicians[index].HOR = [];

          tempResponseData.politicians[index].Comments = "";

          tempResponseErrorStatus.politicians[index].senate = false;
          tempResponseErrorStatus.politicians[index].HOR = false;
        }

        tempResponseData.politicians[index][_key] = value;
        tempResponseData.politicians[index][resetDataOf] =
          typeof tempResponseData.politicians[index][resetDataOf] == "string"
            ? ""
            : [];

        tempResponseErrorStatus.politicians[index][resetDataOf] = false;
      } else {
        let _key: string = key.split("-")[0];
        let resetDataOf: string = key.split("-")[1];

        tempResponseData[_key] = value;
        tempResponseData[resetDataOf] =
          typeof tempResponseData[resetDataOf] == "string" ? "" : [];

        tempResponseErrorStatus[resetDataOf] = false;
      }
    } else if (type == "text") {
      if (section == "Politicians") {
        let index: string = key.split("-")[0];
        let _key: string = key.split("-")[1];

        tempResponseData.politicians[index][_key] = value;
      } else {
        tempResponseData[key] = value;
      }
    } else if (type == "selectAll") {
      tempResponseData[key] = value;
      tempResponseData = selectAllHandler(
        value,
        section,
        tempResponseData,
        _responseData
      );
      tempResponseErrorStatus = resetErrorStatus(
        section,
        tempResponseErrorStatus
      );
    } else if (type == "singleDropdown") {
      let newarr = [];
      newarr.push(value.key);
      tempResponseData[key] = newarr;
      console.log(tempResponseData[key][0]);
    }

    tempResponseData.isEmpty = emptyChecker(tempResponseData);
    tempResponseData = checkForSelectAll(section, tempResponseData);

    setDLResponseData({ ...tempResponseData });
    setDLResponseErrorStatus({ ...tempResponseErrorStatus });
    setDLDropdownOptions({ ...tempDLDropdownOptions });
  };
  const DL_politicianDropDown = (value: string, key: string): IDropDown[] => {
    let _key = key == "HOR" ? "hor" : key;
    let fitleredArr = DLPoliticiansData.filter((_item) => {
      return _item.state == value;
    });
    let _dropDown: IDropDown[] = [];

    fitleredArr.forEach((arr) => {
      if (arr[_key]) {
        if (
          _dropDown.findIndex((optn) => {
            return optn.key == arr[_key];
          }) == -1
        ) {
          _dropDown.push({
            key: arr[_key],
            text: arr[_key],
          });
        }
      }
    });

    return _dropDown;
  };
  const selectAllHandler = (
    condition: boolean,
    section: string,
    curObj: IData,
    initObj: IData
  ): IData => {
    const autoMigrateFunction = (valueKey: string, checkBoxKey: string) => {
      curObj[checkBoxKey] = condition;
      curObj[valueKey] = condition ? curObj[valueKey] : initObj[valueKey];
    };

    if (section == "Schools") {
      autoMigrateFunction("schools", "schoolsChecked");
      autoMigrateFunction("principals", "principalsChecked");
      autoMigrateFunction("instructionCoaches", "instructionCoachesChecked");
      autoMigrateFunction("teachers", "teachersChecked");
      autoMigrateFunction("teachingAssistants", "teachingAssistantsChecked");
      autoMigrateFunction("schoolTeam", "schoolTeamChecked");
      autoMigrateFunction("parents", "parentsChecked");
      autoMigrateFunction("community", "communityChecked");
      autoMigrateFunction("hod", "hodChecked");
      autoMigrateFunction("schoolsOther", "schoolsOtherChecked");
    } else if (section == "Partners") {
      autoMigrateFunction("professionals", "professionalsChecked");
      autoMigrateFunction("services", "servicesChecked");
      autoMigrateFunction("funding", "fundingChecked");
      autoMigrateFunction("resources", "resourcesChecked");
      autoMigrateFunction("governments", "governmentsChecked");
      autoMigrateFunction("indigenous", "indigenousChecked");
      autoMigrateFunction("PA", "PAChecked");
      autoMigrateFunction("media", "mediaChecked");
    }

    return curObj;
  };
  const resetErrorStatus = (
    section: string,
    errorObj: IErrorStatus
  ): IErrorStatus => {
    if (section == "Schools") {
      errorObj.schools = false;
      errorObj.principals = false;
      errorObj.instructionCoaches = false;
      errorObj.teachers = false;
      errorObj.teachingAssistants = false;
      errorObj.schoolTeam = false;
      errorObj.parents = false;
      errorObj.community = false;
      errorObj.hod = false;
      errorObj.schoolsOther = false;
    } else if (section == "Partners") {
      errorObj.professionals = false;
      errorObj.services = false;
      errorObj.funding = false;
      errorObj.resources = false;
      errorObj.governments = false;
      errorObj.indigenous = false;
      errorObj.PA = false;
      errorObj.media = false;
    }

    return errorObj;
  };
  const checkForSelectAll = (section: string, curObj: IData): IData => {
    if (section == "Schools") {
      curObj.schoolsSelectAll =
        curObj.schoolsChecked == true &&
          curObj.principalsChecked == true &&
          curObj.instructionCoachesChecked == true &&
          curObj.teachersChecked == true &&
          curObj.teachingAssistantsChecked == true &&
          curObj.schoolTeamChecked == true &&
          curObj.parentsChecked == true &&
          curObj.communityChecked == true &&
          curObj.hodChecked == true &&
          curObj.schoolsOtherChecked == true
          ? true
          : false;
    } else if (section == "Partners") {
      curObj.partnersSelectAll =
        curObj.professionalsChecked == true &&
          curObj.servicesChecked == true &&
          curObj.fundingChecked == true &&
          curObj.resourcesChecked == true &&
          curObj.governmentsChecked == true &&
          curObj.indigenousChecked == true &&
          curObj.PAChecked == true &&
          curObj.mediaChecked == true
          ? true
          : false;
    }

    return curObj;
  };
  const emptyChecker = (_obj: IData): boolean => {
    let GGSAstatus: boolean = true;
    let Schoolsstatus: boolean = true;
    let DigitalChannelsstatus: boolean = true;
    let Partnersstatus: boolean = true;
    let Politiciansstatus: boolean = true;

    // GGSA
    if (
      _obj.innovationTeamChecked ||
      _obj.employeesChecked ||
      _obj.contractorsChecked ||
      _obj.BAChecked ||
      _obj.managementTeamChecked ||
      _obj.boardChecked ||
      _obj.GGSAComments
    ) {
      GGSAstatus = false;
    }

    // Schools
    if (
      _obj.schoolsChecked ||
      _obj.principalsChecked ||
      _obj.instructionCoachesChecked ||
      _obj.teachersChecked ||
      _obj.teachingAssistantsChecked ||
      _obj.schoolTeamChecked ||
      _obj.parentsChecked ||
      _obj.communityChecked ||
      _obj.hodChecked ||
      _obj.schoolsOtherChecked ||
      _obj.schoolsComments
    ) {
      Schoolsstatus = false;
    }

    // DigitalChannels
    if (
      _obj.gtpChecked ||
      _obj.digitalDatabaseChecked ||
      _obj.websiteChecked ||
      _obj.intranetChecked ||
      _obj.socialMediaChecked ||
      _obj.ceoLinkedInChecked ||
      _obj.productsChecked ||
      _obj.digitalChannelsOtherChecked ||
      _obj.collateralCabinetChecked ||
      _obj.brochureStandChecked ||
      _obj.digitalChannelsComments
    ) {
      DigitalChannelsstatus = false;
    }
    if (
      _obj.professionalsChecked ||
      _obj.servicesChecked ||
      _obj.fundingChecked ||
      _obj.resourcesChecked ||
      _obj.governmentsChecked ||
      _obj.indigenousChecked ||
      _obj.PAChecked ||
      _obj.mediaChecked ||
      _obj.partnersComments
    ) {
      Partnersstatus = false;
    }

    _obj.politicians.forEach((_item: IPoliticianData) => {
      if (
        _item.SATChecked ||
        _item.senateChecked ||
        _item.HORChecked ||
        _item.Comments
      ) {
        Politiciansstatus = Politiciansstatus && false;
      }
    });

    return GGSAstatus == true &&
      Schoolsstatus == true &&
      DigitalChannelsstatus == true &&
      Partnersstatus == true &&
      Politiciansstatus == true
      ? true
      : false;
  };

  const DL_validationFunction = (section: string, nav: string): void => {
    let OverallValidationStatus: boolean = false;

    const validation = (valueKey: string, checkBoxKey: string) => {
      if (DLResponseData[checkBoxKey] && DLResponseData[valueKey].length == 0) {
        DLResponseErrorStatus[valueKey] = true;
        OverallValidationStatus = true;
      }
    };
    const politicianValidation = (
      index: number,
      _key: string,
      checkBoxKey: string
    ) => {
      if (
        DLResponseData.politicians[index][checkBoxKey] &&
        DLResponseData.politicians[index][_key].length == 0
      ) {
        DLResponseErrorStatus.politicians[index][_key] = true;
        OverallValidationStatus = true;
      }
    };

    if (section == "GGSA") {
      validation("innovationTeam", "innovationTeamChecked");
      validation("employees", "employeesChecked");
      validation("contractors", "contractorsChecked");
      validation("BA", "BAChecked");
      validation("managementTeam", "managementTeamChecked");
      validation("board", "boardChecked");
    } else if (section == "Schools") {
      validation("schools", "schoolsChecked");
      validation("principals", "principalsChecked");
      validation("instructionCoaches", "instructionCoachesChecked");
      validation("teachers", "teachersChecked");
      validation("teachingAssistants", "teachingAssistantsChecked");
      validation("schoolTeam", "schoolTeamChecked");
      validation("parents", "parentsChecked");
      validation("community", "communityChecked");
      validation("hod", "hodChecked");
      validation("schoolsOther", "schoolsOtherChecked");
    } else if (section == "DigitalChannels") {
      validation("gtp", "gtpChecked");
      validation("digitalDatabase", "digitalDatabaseChecked");
      validation("website", "websiteChecked");
      validation("intranet", "intranetChecked");
      validation("socialMedia", "socialMediaChecked");
      validation("ceoLinkedIn", "ceoLinkedInChecked");
      validation("products", "productsChecked");
      validation("digitalChannelsOther", "digitalChannelsOtherChecked");
      validation("collateralCabinet", "collateralCabinetChecked");
      validation("brochureStand", "brochureStandChecked");
    } else if (section == "Partners") {
      validation("professionals", "professionalsChecked");
      validation("services", "servicesChecked");
      validation("funding", "fundingChecked");
      validation("resources", "resourcesChecked");
      validation("governments", "governmentsChecked");
      validation("indigenous", "indigenousChecked");
      validation("PA", "PAChecked");
      validation("media", "mediaChecked");
    } else if (section == "Politicians") {
      DLResponseData.politicians.forEach(
        (_item: IPoliticianData, index: number) => {
          politicianValidation(index, "SAT", "SATChecked");
          politicianValidation(index, "senate", "senateChecked");
          politicianValidation(index, "HOR", "HORChecked");
        }
      );
    }
    if (OverallValidationStatus) {
      DLResponseErrorStatus[
        section == "GGSA"
          ? "GGSAOverAllValidation"
          : section == "Schools"
            ? "SchoolsOverAllValidation"
            : section == "DigitalChannels"
              ? "digitalChannelsOverAllValidation"
              : section == "Partners"
                ? "partnersOverAllValidation"
                : section == "Politicians"
                  ? "politiciansOverAllValidation"
                  : null
      ] = OverallValidationStatus;
      setDLLoader("noLoader");
      setDLResponseErrorStatus({ ...DLResponseErrorStatus });
    } else {
      nav == "Submit" || nav == "SendForApproval" || nav == "PRSubmit"
        ? DLResponseData.isEmpty
          ? (setDLLoader("noLoader"), setDLWarningPopup(true))
          : DLResponseData.action == "new"
            ? addDLFunction(nav)
            : updateDLFunction(nav)
        : setDLPageSwitch(nav);
    }
  };

  const DL_AccessListGetItems = (): void => {
    _webURL.lists
      .getByTitle(DLAccessListName)
      // .items.select(
      //   "*")
      .items.filter(`UsersId eq ${loggedUserID}`)
      .top(5000)
      .get()
      .then((items) => {
        // console.log(items, "List of Role's access")
        let _perm: string = "";
        if (items.length > 0) {
          let allPerms: string[] = [];
          items.forEach((_item) => {
            allPerms.push(_item.PermissionType);

            if (allPerms.length == items.length) {
              if (allPerms.some((perm) => perm == "All Access")) {
              console.log("FullPerms")
              _perm = "FullPerms";
                  } else if (allPerms.some((perm) => perm == "High Level")) {
                    // console.log("HighPerms")
                    _perm = "HighPerms";
                  } else if (allPerms.some((perm) => perm == "Medium Level")) {
                    console.log("MediumPerms")
              _perm = "MediumPerms";
                  }
            }
          });
        } else {
          _perm = "LowPerms";
        }

        setDLUserPermission(_perm);

        // {DLPageSwitch == "GGSA"
        // ? GGSA_Tab()
        // : DLPageSwitch == "Schools"
        //   ? Schools_Tab()
        //   : DLPageSwitch == "DigitalChannels"


        setDLPageSwitch("GGSA");
        // setDLPageSwitch(_perm == "MediumPerms" ? "GGSA":"Schools" );
        DL_getPoliticiansData();
      })
      .catch((err) => {
        DL_ErrorFunction(err, "DL_AccessListGetItems");
      });
  };
  const DL_getPoliticiansData = (): void => {
    _webURL.lists
      .getByTitle("Politicians")
      .items.get()
      .then((_plData: any) => {
        let plData: IAllPoliticiansData[] = [];
        _plData.forEach((pl: any) => {
          if (pl.Title) {
            plData.push({
              state: pl.Title,
              senate: pl.Senate,
              hor: pl.HouseOfRepresentatives,
            });
          }
        });

        DL_getAllDropdownOptions([...plData]);
      })
      .catch((err) => {
        DL_ErrorFunction(err, "DL_getPoliticiansData");
      });
  };
  const DL_getAllDropdownOptions = (
    _politiciansData: IAllPoliticiansData[]
  ): void => {
    const getDropDownFunction = (
      listName: string,
      optnKey: string,
      key: string,
      text: string,
      filter?: string
    ) => {
      let multipleTextCheck: boolean = false;
      let text1: string = "";
      let text2: string = "";

      if (text.includes("-")) {
        multipleTextCheck = true;
        text1 = text.split("-")[0];
        text2 = text.split("-")[1];
      }

      _webURL.lists
        .getByTitle(listName)
        .items.filter(filter ? filter : "")
        .top(5000)
        .get()
        .then((response) => {
          response.forEach((res) => {
            if (res[key]) {
              if (
                DLDropdownOptions[optnKey].findIndex((optn) => {
                  return optn.key == res[key];
                }) == -1
              ) {
                DLDropdownOptions[optnKey].push({
                  key: res[key],
                  text: multipleTextCheck
                    ? `${res[text1]} ${res[text2]}`
                    : res[text],
                });
              }
            }
            if (DLDropdownOptions[optnKey].length == 1) {
              DLDropdownOptions[optnKey].unshift({
                key: "All",
                text: "All",
              });
            }
          });
        });
    };

    const getYesOrNoChoices = (optnKey: string) => {
      DLDropdownOptions[optnKey].push(
        {
          key: "Yes",
          text: "Yes",
        },
        {
          key: "No",
          text: "No",
        }
      );
    };

    getDropDownFunction(
      "InnovationTeam",
      "innovationTeamOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Employees",
      "employeesOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Contractors",
      "contractorsOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Management",
      "managementTeamOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction("Board", "boardOptns", "Email", "FirstName-LastName");

    getDropDownFunction(
      "Schools List",
      "principalsOptns",
      "Email",
      "FirstName-LastName",
      "Role eq 'Principal'"
    );
    getDropDownFunction(
      "InstructionCoaches",
      "instructionCoachesOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Teachers",
      "teachersOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "TeachingAssistants",
      "teachingAssistantsOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction("Schools List", "schoolsOptns", "School", "School");
    getDropDownFunction(
      "SchoolTeams",
      "schoolTeamOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Parents",
      "parentsOptns",
      "School",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Communities",
      "communityOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "HeadOfDepartments",
      "hodOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Other",
      "schoolsOtherOptns",
      "Email",
      "FirstName-LastName"
    );

    getDropDownFunction(
      "Master Product List",
      "productsOptns",
      "Title",
      "Title"
    );
    getDropDownFunction(
      "Other",
      "digitalChannelsOtherOptns",
      "Email",
      "FirstName-LastName"
    );

    getDropDownFunction(
      "Professionals",
      "professionalsOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Services",
      "servicesOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Funding",
      "fundingOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Resources",
      "resourcesOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Governments",
      "governmentsOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Indigenous",
      "indigenousOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction(
      "Professional Associations",
      "PAOptns",
      "Email",
      "FirstName-LastName"
    );
    getDropDownFunction("Media", "mediaOptns", "Email", "FirstName-LastName");

    getYesOrNoChoices("gtpOptns");
    getYesOrNoChoices("digitalDatabaseOptns");
    getYesOrNoChoices("websiteOptns");
    getYesOrNoChoices("intranetOptns");
    getYesOrNoChoices("socialMediaOptns");
    getYesOrNoChoices("ceoLinkedInOptns");
    getYesOrNoChoices("collateralCabinetOptns");
    getYesOrNoChoices("brochureStandOptns");

    // BusinessArea
    _webURL.lists
      .getByTitle(DLListName)
      .fields.getByInternalNameOrTitle("BusinessArea")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              DLDropdownOptions.BAOptns.findIndex((baOptn) => {
                return baOptn.key == choice;
              }) == -1
            ) {
              DLDropdownOptions.BAOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
          if (DLDropdownOptions.BAOptns.length == 1) {
            DLDropdownOptions.BAOptns.unshift({
              key: "All",
              text: "All",
            });
          }
        });
      });

    // State and Territory
    _politiciansData.forEach((optn) => {
      if (optn.state) {
        if (
          DLDropdownOptions.SATOptns.findIndex((_optn) => {
            return _optn.key == optn.state;
          }) == -1
        ) {
          DLDropdownOptions.SATOptns.push({
            key: optn.state,
            text: optn.state,
          });
        }
      }
    });

    setDLPoliticiansData([..._politiciansData]);
    setDLDropdownOptions({ ...DLDropdownOptions });

    DL_GetItemsFunction(props.docReviewDetails.dlID);
  };
  const DL_GetItemsFunction = (dlID: number): void => {
    let tempDLResponseData: IData = { ...DLResponseData };
    let tempDLResponseErrorStatus: IErrorStatus = {
      ...DLResponseErrorStatus,
    };
    if (dlID != null) {
      _webURL.lists
        .getByTitle(DLListName)
        .items.filter(`ID eq ${dlID}`)
        .top(5000)
        .orderBy("Modified", false)
        .get()
        .then((items) => {
          let item: any = items[0];

          editableStatus =
            props.docReviewDetails.isEditable == true &&
              item.ApprovalStatus != "Approved"
              ? true
              : false;
          const autoMigrateFunction = (key: string) => {
            return item[key] ? item[key].split(";") : [];
          };

          let politiciansJSON: IPoliticianData[];

          if (item.PoliticiansJSON) {
            politiciansJSON = JSON.parse(item.PoliticiansJSON);
            politiciansJSON.push({
              SATChecked: false,
              SAT: "",
              senateChecked: false,
              senate: [],
              HORChecked: false,
              HOR: [],
              Comments: "",
            });
          } else {
            politiciansJSON = [
              {
                SATChecked: false,
                SAT: "",
                senateChecked: false,
                senate: [],
                HORChecked: false,
                HOR: [],
                Comments: "",
              },
            ];
          }

          tempDLResponseData = {
            isEmpty: false,
            ID: item.ID,
            action: "edit",
            // GGSA
            innovationTeamChecked: item.InnovationTeam ? true : false,
            innovationTeam: autoMigrateFunction("InnovationTeam"),
            employeesChecked: item.Employees ? true : false,
            employees: autoMigrateFunction("Employees"),

            contractorsChecked: item.Contractors ? true : false,
            contractors: autoMigrateFunction("Contractors"),
            BAChecked:
              item.BusinessArea != null && item.BusinessArea.length > 0
                ? true
                : false,
            BA:
              item.BusinessArea != null && item.BusinessArea.length > 0
                ? [...item.BusinessArea]
                : [],
            managementTeamChecked: item.ManagementTeam ? true : false,
            managementTeam: autoMigrateFunction("ManagementTeam"),
            boardChecked: item.Board ? true : false,
            board: autoMigrateFunction("Board"),
            GGSAComments: item.ggsaComments,

            // Schools
            schoolsSelectAll: false,

            schoolsChecked: item.Schools ? true : false,
            schools: autoMigrateFunction("Schools"),
            principalsChecked: item.Principals ? true : false,
            principals: autoMigrateFunction("Principals"),
            instructionCoachesChecked: item.InstructionCoaches ? true : false,
            instructionCoaches: autoMigrateFunction("InstructionCoaches"),
            teachersChecked: item.Teachers ? true : false,
            teachers: autoMigrateFunction("Teachers"),
            teachingAssistantsChecked: item.TeachingAssistant ? true : false,
            teachingAssistants: autoMigrateFunction("TeachingAssistant"),
            schoolTeamChecked: item.SchoolTeams ? true : false,
            schoolTeam: autoMigrateFunction("SchoolTeams"),
            parentsChecked: item.Parents ? true : false,
            parents: autoMigrateFunction("Parents"),
            communityChecked: item.Communities ? true : false,
            community: autoMigrateFunction("Communities"),
            hodChecked: item.HODs ? true : false,
            hod: autoMigrateFunction("HODs"),
            schoolsOtherChecked: item.SchoolOthers ? true : false,
            schoolsOther: autoMigrateFunction("SchoolOthers"),
            schoolsComments: item.SchoolsComments,

            // Digital Channels
            gtpChecked: item.GreatTeachingPortal ? true : false,
            gtp: autoMigrateFunction("GreatTeachingPortal"),
            digitalDatabaseChecked: item.DigitalDatabase ? true : false,
            digitalDatabase: autoMigrateFunction("DigitalDatabase"),
            websiteChecked: item.Website ? true : false,
            website: autoMigrateFunction("Website"),
            intranetChecked: item.Intranet ? true : false,
            intranet: autoMigrateFunction("Intranet"),
            brochureStandChecked: item.BrochureStand ? true : false,
            brochureStand: autoMigrateFunction("BrochureStand"),
            collateralCabinetChecked: item.CollateralCabinet ? true : false,
            collateralCabinet: autoMigrateFunction("CollateralCabinet"),
            socialMediaChecked: item.SocialMedia ? true : false,
            socialMedia: autoMigrateFunction("SocialMedia"),
            ceoLinkedInChecked: item.CEOLinkedIn ? true : false,
            ceoLinkedIn: autoMigrateFunction("CEOLinkedIn"),
            productsChecked: item.Products ? true : false,
            products: autoMigrateFunction("Products"),
            digitalChannelsOtherChecked: item.DCOthers ? true : false,
            digitalChannelsOther: autoMigrateFunction("DCOthers"),
            digitalChannelsComments: item.DCComments,

            // Partners
            partnersSelectAll: false,

            professionalsChecked: item.Professionals ? true : false,
            professionals: autoMigrateFunction("Professionals"),
            servicesChecked: item.Services ? true : false,
            services: autoMigrateFunction("Services"),
            fundingChecked: item.Funding ? true : false,
            funding: autoMigrateFunction("Funding"),
            resourcesChecked: item.Resources ? true : false,
            resources: autoMigrateFunction("Resources"),
            governmentsChecked: item.Governments ? true : false,
            governments: autoMigrateFunction("Governments"),
            indigenousChecked: item.Indigenous ? true : false,
            indigenous: autoMigrateFunction("Indigenous"),
            PAChecked: item.ProfessionalAssociations ? true : false,
            PA: autoMigrateFunction("ProfessionalAssociations"),
            mediaChecked: item.Media ? true : false,
            media: autoMigrateFunction("Media"),
            partnersComments: item.PartnersComments,

            //Politicians
            politiciansSelectAll: item.PoliticiansJSON ? true : false,
            politicians: politiciansJSON,
          };

          tempDLResponseData.isEmpty = emptyChecker(tempDLResponseData);
          tempDLResponseData = checkForSelectAll("Schools", tempDLResponseData);
          tempDLResponseData = checkForSelectAll(
            "Partners",
            tempDLResponseData
          );
          if (tempDLResponseData.politicians.length > 1) {
            tempDLResponseData.politicians.forEach(() => {
              tempDLResponseErrorStatus.politicians.push({
                SAT: false,
                senate: false,
                HOR: false,
              });
            });
          }

          _webURL.lists
            .getByTitle(DRListName)
            .items.filter(`DistributionListID eq ${dlID}`)
            .top(5000)
            .orderBy("Modified", false)
            .get()
            .then((_drData) => {
              let filteredDRData = _drData.filter((drData) => {
                return (
                  drData.auditRequestType == "Distribute" &&
                  drData.auditResponseType == "Distribute pending"
                );
              });
              filteredDRData.length > 0
                ? setDisableApproveBtn(true)
                : setDisableApproveBtn(false);
              setDLResponseData({ ...tempDLResponseData });
              setDLResponseErrorStatus({ ...tempDLResponseErrorStatus });
              setDLLoader("noLoader");
            })
            .catch((err) => {
              DL_ErrorFunction(err, "DL_GetItemsFunction-getDRItems");
            });
        })
        .catch((err) => {
          DL_ErrorFunction(err, "DL_GetItemsFunction");
        });
    } else {
      editableStatus = props.docReviewDetails.isEditable == true ? true : false;
      tempDLResponseData.action = "new";

      setDLResponseData({ ...tempDLResponseData });
      setDLResponseErrorStatus({ ...tempDLResponseErrorStatus });
      setDLLoader("noLoader");
    }
  };

  const responseDataGenerator = (_responseData: IData) => {
    const dataModifierFunction = (key: string): string => {
      return _responseData[key].length > 0 ? _responseData[key].join(";") : "";
    };
    const politiciansJSONFunction = (
      politiciansJSONFunction: IPoliticianData[]
    ): string => {
      if (
        politiciansJSONFunction.length == 1 &&
        !politiciansJSONFunction[0].SATChecked
      ) {
        return "";
      } else {
        let filteredPoliticianJSON: IPoliticianData[] =
          politiciansJSONFunction.filter((_arr: IPoliticianData) => {
            return _arr.SATChecked;
          });

        return JSON.stringify(filteredPoliticianJSON);
      }
    };

    let responseData = {
      Title: props.spcontext.pageContext.user.email
        ? props.spcontext.pageContext.user.email
        : "",
      // GGSA
      BusinessArea:
        DLResponseData.BA.length > 0
          ? { results: [...DLResponseData.BA] }
          : { results: [] },
      InnovationTeam: dataModifierFunction("innovationTeam"),
      Employees: dataModifierFunction("employees"),
      Contractors: dataModifierFunction("contractors"),
      ManagementTeam: dataModifierFunction("managementTeam"),
      Board: dataModifierFunction("board"),
      ggsaComments: DLResponseData.GGSAComments,

      // Schools
      Principals: dataModifierFunction("principals"),
      InstructionCoaches: dataModifierFunction("instructionCoaches"),
      Teachers: dataModifierFunction("teachers"),
      TeachingAssistant: dataModifierFunction("teachingAssistants"),
      Schools: dataModifierFunction("schools"),
      SchoolTeams: dataModifierFunction("schoolTeam"),
      Parents: dataModifierFunction("parents"),
      Communities: dataModifierFunction("community"),
      HODs: dataModifierFunction("hod"),
      SchoolOthers: dataModifierFunction("schoolsOther"),
      SchoolsComments: DLResponseData.schoolsComments,

      // DigitalChannels
      GreatTeachingPortal: dataModifierFunction("gtp"),
      DigitalDatabase: dataModifierFunction("digitalDatabase"),
      Website: dataModifierFunction("website"),
      Intranet: dataModifierFunction("intranet"),
      SocialMedia: dataModifierFunction("socialMedia"),
      CEOLinkedIn: dataModifierFunction("ceoLinkedIn"),
      Products: dataModifierFunction("products"),
      DCOthers: dataModifierFunction("digitalChannelsOther"),
      CollateralCabinet: dataModifierFunction("collateralCabinet"),
      BrochureStand: dataModifierFunction("brochureStand"),
      DCComments: DLResponseData.digitalChannelsComments,

      // Partners
      Professionals: dataModifierFunction("professionals"),
      Services: dataModifierFunction("services"),
      Funding: dataModifierFunction("funding"),
      Resources: dataModifierFunction("resources"),
      Governments: dataModifierFunction("governments"),
      Indigenous: dataModifierFunction("indigenous"),
      ProfessionalAssociations: dataModifierFunction("PA"),
      Media: dataModifierFunction("media"),
      PartnersComments: DLResponseData.partnersComments,

      //Politicians 
      PoliticiansJSON: politiciansJSONFunction(DLResponseData.politicians),

      ApprovalStatus: "Pending",
      // props.docReviewDetails.response == "Publish ready"
      //   ? "Approved"
      //   : "Pending",
    };

    return responseData;
  };
  const addDLFunction = async (nav: string): Promise<void> => {
    let responseData = responseDataGenerator(DLResponseData);
    let approverEmail: string = DL_getApprover(
      DLApproverData.dlConfigData,
      DLApproverData.userData
    );

    if (
      nav == "PRSubmit" ||
      ((nav == "Submit" || nav == "SendForApproval") && approverEmail == "")
    ) {
      responseData.ApprovalStatus = "Approved";
    }
    await _webURL.lists
      .getByTitle(DLListName)
      .items.add(responseData)
      .then(async (e) => {
        let targetID: number = e.data.ID;

        if (props.docReviewDetails.response == "Signed Off") {
          if (nav == "Submit") {
            await updateDRFunction(targetID);
          } else if (nav == "SendForApproval") {
            approverEmail ? await addDRFunction(targetID, approverEmail) : null;
            await updateDRFunction(targetID);
          }
          // await addDRFunction(targetID);
          // await updateDRFunction(targetID);
        } else {
          await updateDRFunction(targetID);
        }

        await props.distributionListHandlerFunction("DocumentReview", null);
      })
      .catch((err) => {
        DL_ErrorFunction(err, "addDLFunction");
      });
  };
  const updateDLFunction = async (nav: string): Promise<void> => {
    let responseData = responseDataGenerator(DLResponseData);
    let approverEmail: string = DL_getApprover(
      DLApproverData.dlConfigData,
      DLApproverData.userData
    );

    if (
      nav == "PRSubmit" ||
      ((nav == "Submit" || nav == "SendForApproval") && approverEmail == "")
    ) {
      responseData.ApprovalStatus = "Approved";
    }
    await _webURL.lists
      .getByTitle(DLListName)
      .items.getById(DLResponseData.ID)
      .update(responseData)
      .then(async () => {
        if (nav == "SendForApproval") {
          approverEmail
            ? await addDRFunction(DLResponseData.ID, approverEmail)
            : null;
        }

        await props.distributionListHandlerFunction("DocumentReview", null);
      })
      .catch((err) => {
        DL_ErrorFunction(err, "updateDLFunction");
      });
  };
  const addDRFunction = async (
    dlID: number,
    _approverEmail: string
  ): Promise<void> => {
    await _webURL.lists
      .getByTitle(DRListName)
      .items.getById(props.docReviewDetails.docReviewID)
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
      .expand("FromUser,ToUser,CcEmail")
      .get()
      .then(async (item) => {
        let approverDetails = props.peopleList.filter((obj) => {
          return obj.secondaryText == _approverEmail;
        });
        const responseData = {
          Title: item.Title ? item.Title : "",
          auditLink: item.auditLink ? item.auditLink : "",
          auditRequestType: "Distribute",
          auditSent: moment().format("MM/DD/yyyy"),
          CcEmailId: item.CcEmailId
            ? { results: item.CcEmailId }
            : { results: [] },
          auditResponseType: "Distribute pending",
          FeedbackRepeated: false,
          auditFrom: item.auditFrom ? item.auditFrom : "",
          FromEmail: item.FromEmail ? item.FromEmail : "",
          FromUserId: item.FromUserId ? item.FromUserId : "",

          auditTo: approverDetails.length > 0 ? approverDetails[0].text : "",
          ToEmail: approverDetails.length > 0 ? _approverEmail : "",
          ToUserId: approverDetails.length > 0 ? approverDetails[0].ID : "",

          auditComments: item.auditComments ? item.auditComments : "",
          auditDocLink: item.auditDocLink ? item.auditDocLink : "",
          auditDepartment: item.auditDepartment ? item.auditDepartment : "",
          auditLastResponse: item.auditLastResponse
            ? item.auditLastResponse
            : "",
          auditID: item.auditID ? item.auditID : "",
          auditReplyTo: item.auditReplyTo ? item.auditReplyTo : "",
          auditResponseMeetingRequired: false,
          ResponseAcknowledged: false,
          AnnualPlanID: item.AnnualPlanID ? item.AnnualPlanID : null,
          DeliveryPlanID: item.DeliveryPlanID ? item.DeliveryPlanID : null,
          ProductionBoardID: item.ProductionBoardID
            ? item.ProductionBoardID
            : null,
          DRPageName: item.DRPageName ? item.DRPageName : null,
          DistributionListID: dlID,
        };

        await _webURL.lists
          .getByTitle("Review Log")
          .items.add(responseData)
          .then((_item) => { })
          .catch(async (err) => {
            await DL_ErrorFunction(err, "addDRFunction-addItem");
          });
      })
      .catch((err) => {
        DL_ErrorFunction(err, "addDRFunction-getItem");
      });
  };
  const updateDRFunction = async (dlID: number): Promise<void> => {
    await _webURL.lists
      .getByTitle(DRListName)
      .items.getById(props.docReviewDetails.docReviewID)
      .update({ DistributionListID: dlID })
      .then(async () => {
        await [];
      })
      .catch((err) => {
        DL_ErrorFunction(err, "updateDRFunction");
      });
  };

  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  const DL_ErrorFunction = (err: string, functionName: string): void => {
    console.log(err, functionName);

    let response = {
      ComponentName: "Distribution list",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(err["message"]),
      Title: loggedUserEmail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setDLLoader("noLoader");
        ErrorPopup();
      }
    );
  };

  const DL_EmailTemplateToGenerator = (): string => {
    let initialArr: string[] = [];
    let filteredArr: string[] = [];
    let finalStr: string = "";

    const loopingFunction = (key: string): void => {
      if (DLResponseData[key].length > 0) {
        DLResponseData[key].forEach((option: string) => {
          if (option != "All") {
            initialArr.push(option);
          }
        });
      }
    };

    loopingFunction("innovationTeam");
    loopingFunction("employees");
    loopingFunction("contractors");
    loopingFunction("managementTeam");
    loopingFunction("board");
    loopingFunction("principals");
    loopingFunction("instructionCoaches");
    loopingFunction("teachers");
    loopingFunction("teachingAssistants");
    loopingFunction("schoolTeam");
    loopingFunction("community");
    loopingFunction("hod");
    loopingFunction("schoolsOther");
    loopingFunction("digitalChannelsOther");
    loopingFunction("professionals");
    loopingFunction("services");
    loopingFunction("funding");
    loopingFunction("resources");
    loopingFunction("governments");
    loopingFunction("indigenous");
    loopingFunction("PA");
    loopingFunction("media");

    console.log(initialArr);

    initialArr.forEach((res: string) => {
      if (
        res &&
        filteredArr.findIndex((optn) => {
          return optn == res;
        }) == -1
      ) {
        filteredArr.push(res);
      }
    });
    console.log(filteredArr);

    finalStr = filteredArr.join(";");
    finalStr = finalStr ? finalStr + ";" : "";
    console.log(finalStr);

    return finalStr;
  };

  const DL_getMasterUserData = (): void => {
    _webURL.lists
      .getByTitle(MasterUserListName)
      .items.filter(`UserId eq ${loggedUserID}`)
      .top(5000)
      .get()
      .then((mulData) => {
        mulData.length > 0 ? DL_getApprovalConfigData(mulData[0]) : null;
      })
      .catch((err) => {
        DL_ErrorFunction(err, "DL_getMasterUserData");
      });
  };
  const DL_getApprovalConfigData = (_userData: any): void => {
    _webURL.lists
      .getByTitle(DLApprovalConfigListName)
      .items.filter(`BusinessArea eq '${_userData.BusinessArea}'`)
      .select("*", "To/Title", "To/Id", "To/EMail")
      .expand("To")
      .top(5000)
      .get()
      .then((items: any) => {
        items.length > 0
          ? setDLApproverData({ dlConfigData: items[0], userData: _userData })
          : setDLApproverData(null);
        DL_AccessListGetItems();
      })
      .catch((err) => {
        DL_ErrorFunction(err, "DL_getApprovalConfigData");
      });
  };
  const DL_getApprover = (DACData: any, userData: any): string => {
    const approverReturnFunction = (key: string, _key: string): void => {
      // if (DLResponseData[key].length > 0) {
      let _condition: boolean =
        _key != "Other"
          ? DLResponseData[key].length > 0
          : DLResponseData.schoolsOther.length > 0 ||
          DLResponseData.digitalChannelsOther.length > 0;

      if (_condition) {
        let approvalConfigArr = DACData[_key].split("~");

        if (approvalConfigArr[0] && approvalConfigArr[0] == "Mandatory") {
          approverEmail = DACData.ToId ? DACData.To.EMail : "";
        } else if (
          approvalConfigArr[1] &&
          approvalConfigArr[1].split(",").length > 0 &&
          !approvalConfigArr[1]
            .split(",")
            .some((_pos) => _pos == userData.Position)
        ) {
          approverEmail = DACData.ToId ? DACData.To.EMail : "";
        } else if (
          approvalConfigArr[2] &&
          approvalConfigArr[2].split(",").length > 0 &&
          !approvalConfigArr[2]
            .split(",")
            .some((_people) => _people == loggedUserEmail)
        ) {
          approverEmail = DACData.ToId ? DACData.To.EMail : "";
        }
      }
    };

    let approverEmail: string = "";
    approverReturnFunction("employees", "Employees");
    approverReturnFunction("contractors", "Contractors");
    approverReturnFunction("managementTeam", "Management");
    approverReturnFunction("board", "Board");
    approverReturnFunction("schools", "Schools");
    approverReturnFunction("instructionCoaches", "InstructionCoaches");
    approverReturnFunction("teachers", "Teachers");
    approverReturnFunction("teachingAssistants", "TeachingAssistants");
    approverReturnFunction("schoolTeam", "SchoolTeams");
    approverReturnFunction("parents", "Parents");
    approverReturnFunction("community", "Communities");
    approverReturnFunction("", "Other");
    approverReturnFunction("professionals", "Professionals");
    approverReturnFunction("services", "Services");
    approverReturnFunction("funding", "Funding");
    approverReturnFunction("resources", "Resources");
    approverReturnFunction("governments", "Governments");
    approverReturnFunction("politicians", "Politicians");
    approverReturnFunction("indigenous", "Indigenous");
    // approverReturnFunction("", "ProfessionalAssociations");
    approverReturnFunction("media", "Media");

    return approverEmail;
  };
  // Function-Declaration Ends

  useEffect(() => {
    setDLLoader("StartLoader");
    DL_getMasterUserData();
  }, [DLReRender]);
  return (
    <div style={{ padding: "5px 15px" }}>
      {DLLoader == "StartLoader" ? (
        <CustomLoader />
      ) : (
        <div>
          {/* header-Section Starts*/}
          <div style={{ marginBottom: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  marginBottom: 10,
                }}
              >
                <Icon
                  aria-label="ChevronLeftMed"
                  iconName="NavigateBack"
                  className={DLIconStyleClass.backIcon}
                  onClick={() => {
                    props.distributionListHandlerFunction(
                      "DocumentReview",
                      null
                    );
                  }}
                />
                <Label styles={headingStyles}>Distribution</Label>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
              >
                <button
                  style={{
                    border: "none",
                    color: "#fff",
                    backgroundColor: "#faa332",
                    fontSize: 12,
                    fontWeight: 600,
                    borderRadius: 5,
                    marginRight: 10,
                    padding: "5px 20px",
                    cursor: "pointer",
                  }}
                  onClick={() => {
                    // setMailTemplatePopup(true);

                    const emailURL = `mailto:${DL_EmailTemplateToGenerator()}?cc=''&subject=''&body=
                    ${encodeURIComponent(`The development and support sentences act as the body of the paragraph. Development sentences elaborate and explain the idea
                    with details too specific for the topic sentence, while support
                    sentences provide evidence, opinions, or other statements that
                    back up or confirm the paragraph’s main idea. Last, the conclusion
                    wraps up the idea, sometimes summarizing what’s been presented or
                    transitioning to the next paragraph. The content of the conclusion
                    depends on the type of paragraph, and it’s often acceptable to end
                    a paragraph with a final piece of support that concludes the
                    thought instead of a summary`)}`;
                    window.location.href = emailURL;

 
                  }}
                >
                Open mail template
              </button>
            </div>
          </div>
          <div className={styles.dlHeaderButtonSection}>
            {DLUserPermission == "FullPerms" ||
              DLUserPermission == "HighPerms" ||
              DLUserPermission == "MediumPerms" ? (
              <button
                className={
                  DLPageSwitch == "GGSA"
                    ? styles.activeButton
                    : styles.inactiveButton
                }
                onClick={() => {
                  DL_validationFunction(DLPageSwitch, "GGSA");
                }}
              >
                GGSA
              </button>
            ) : null}
            {DLUserPermission == "FullPerms" ||
              DLUserPermission == "HighPerms" ? (
              <button
                className={
                  DLPageSwitch == "Schools"
                    ? styles.activeButton
                    : styles.inactiveButton
                }
                onClick={() => {
                  DL_validationFunction(DLPageSwitch, "Schools");
                }}
              >
                Schools
              </button>
            ) : null}
            {DLUserPermission == "FullPerms" ||
              DLUserPermission == "HighPerms" ||
              DLUserPermission == "MediumPerms" ? (
              <button
                className={
                  DLPageSwitch == "DigitalChannels"
                    ? styles.activeButton
                    : styles.inactiveButton
                }
                onClick={() => {
                  DL_validationFunction(DLPageSwitch, "DigitalChannels");
                }}
              >
                Digital channels
              </button>
            ) : null}
            {DLUserPermission == "FullPerms" ||
              DLUserPermission == "HighPerms" ? (
              <button
                className={
                  DLPageSwitch == "Partners"
                    ? styles.activeButton
                    : styles.inactiveButton
                }
                onClick={() => {
                  DL_validationFunction(DLPageSwitch, "Partners");
                }}
              >
                Partners
              </button>
            ) : null}
            {DLUserPermission == "FullPerms" ? (
              <button
                className={
                  DLPageSwitch == "Politicians"
                    ? styles.activeButton
                    : styles.inactiveButton
                }
                onClick={() => {
                  DL_validationFunction(DLPageSwitch, "Politicians");
                }}
              >
                Politicians
              </button>
            ) : null}
          </div>
        </div>
          {/* header-Section Ends*/}
      {/* body-Section Starts */}
      <div>
        {DLPageSwitch == "GGSA"
          ? GGSA_Tab()
          : DLPageSwitch == "Schools"
            ? Schools_Tab()
            : DLPageSwitch == "DigitalChannels"
              ? DigitalChannels_Tab()
              : DLPageSwitch == "Partners"
                ? Partners_Tab()
                : DLPageSwitch == "Politicians"
                  ? Politicians_Tab()
                  : null}
      </div>
      {/* body-Section Ends */}
      {/* Popup-Section Starts */}
      {DLWarningPopup ? (
        <Modal
          isOpen={DLWarningPopup}
          isBlocking={true}
          styles={DLModalStyles}
        >
          <div>
            <Label className={styles.dlPopupLabel}>Warning</Label>
            <div className={styles.dlPopupDescription}>
              Empty form cannot be submitted.
            </div>
            <div className={styles.dlPopupButtonSection}>
              <button
                onClick={() => {
                  setDLWarningPopup(false);
                }}
              >
                Close
              </button>
            </div>
          </div>
        </Modal>
      ) : null}
      {mailTemplatePopup ? (
        <Modal isOpen={mailTemplatePopup} isBlocking={true}>
          <div style={{ padding: "30px 20px" }}>
            <Label
              className={styles.atpPopupLabel}
              style={{ marginLeft: "auto" }}
            >
              Distribute
            </Label>

            <div
              style={{
                maxHeight: 85,
                overflow: "auto",
                width: 800,
                paddingRight: 20,
                overflowWrap: "break-word",
                marginBottom: 20,
              }}
            >
              <Label>
                To
                <Icon
                  iconName="Copy"
                  style={{
                    color: "#2392B2",
                    marginLeft: 7,
                    cursor: "pointer",
                  }}
                  onClick={() => {
                    navigator.clipboard.writeText(
                      DL_EmailTemplateToGenerator()
                    );
                  }}
                />
              </Label>
              <>{DL_EmailTemplateToGenerator()}</>
            </div>
            <div
              style={{
                width: 800,
                paddingRight: 20,
                overflowWrap: "break-word",
                marginBottom: 20,
              }}
            >
              <Label>
                CC
                <Icon
                  iconName="Copy"
                  style={{
                    color: "#2392B2",
                    marginLeft: 7,
                    cursor: "pointer",
                  }}
                  onClick={() => {
                    // navigator.clipboard.writeText(loggedUserEmail);
                  }}
                />
              </Label>
              <></>
              {/* <>{loggedUserEmail}</> */}
            </div>
            <div
              style={{
                width: 800,
                paddingRight: 20,
                overflowWrap: "break-word",
                marginBottom: 20,
              }}
            >
              <Label>
                Subject
                <Icon
                  iconName="Copy"
                  style={{
                    color: "#2392B2",
                    marginLeft: 7,
                    cursor: "pointer",
                  }}
                  onClick={() => {
                    // navigator.clipboard.writeText(
                    //   "Reg:Paragraph structured"
                    // );
                  }}
                />
              </Label>
              <></>
              {/* <>Reg:Paragraph structured</> */}
            </div>
            <div
              style={{
                width: "800px",
                paddingRight: "20px",
                overflowWrap: "break-word",
                marginBottom: 20,
              }}
            >
              <Label>
                Body
                <Icon
                  iconName="Copy"
                  style={{
                    color: "#2392B2",
                    marginLeft: 7,
                    cursor: "pointer",
                  }}
                  onClick={() => {
                    navigator.clipboard.writeText(
                      `The development and support sentences act as the body of the paragraph. Development sentences elaborate and explain the idea
                    with details too specific for the topic sentence, while support
                    sentences provide evidence, opinions, or other statements that
                    back up or confirm the paragraph’s main idea. Last, the conclusion
                    wraps up the idea, sometimes summarizing what’s been presented or
                    transitioning to the next paragraph. The content of the conclusion
                    depends on the type of paragraph, and it’s often acceptable to end
                    a paragraph with a final piece of support that concludes the
                    thought instead of a summary`
                    );
                  }}
                />
              </Label>
              <>
                The development and support sentences act as the body of the
                paragraph. Development sentences elaborate and explain the
                idea with details too specific for the topic sentence, while
                support sentences provide evidence, opinions, or other
                statements that back up or confirm the paragraph’s main
                idea. Last, the conclusion wraps up the idea, sometimes
                summarizing what’s been presented or transitioning to the
                next paragraph. The content of the conclusion depends on the
                type of paragraph, and it’s often acceptable to end a
                paragraph with a final piece of support that concludes the
                thought instead of a summary.{" "}
              </>
            </div>
            <div
              style={{
                display: "flex",
                justifyContent: "flex-end",
                padding: "10px 0px",
              }}
            >
              <button
                style={{
                  border: "none",
                  color: "#fff",
                  backgroundColor: "#b80000",
                  fontSize: 12,
                  fontWeight: 600,
                  borderRadius: 5,
                  marginRight: 10,
                  padding: "5px 20px",
                  cursor: "pointer",
                }}
                onClick={() => {
                  setMailTemplatePopup(false);
                }}
              >
                Close
              </button>
            </div>
          </div>
        </Modal>
      ) : null}
      {/* Popup-Section Ends */}
    </div>
  )
}
    </div >
  );
};

export default DistributionList;
