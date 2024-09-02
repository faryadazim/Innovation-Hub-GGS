import * as React from "react";
import { useState, useEffect } from "react";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import AnnualPlan from "./AnnualPlan";
import DeliveryPlan from "./DeliveryPlan";
import ProductionBoard from "./ProductionBoard";
import DocumentReview from "./DocumentReview";
import ActivityPlan from "./ActivityPlan";
import ActivityDeliveryPlan from "./ActivityDeliveryPlan2";
import ActivityTemplate from "./ActivityTemplate";
import ActivityProductionBoard from "./ActivityProductionBoard";
import PL_MainComponent from "./PL_MainComponent";
import MasterProduct from "./MasterProduct";
import DocumentReviewHistory from "./DocumentReviewHistory";
import OrgReporting from "./OrgReporting";
import StockList from "./StockList";
import StaffList from "./StaffList";
import DistributionApprovalConfig from "./DistributionApprovalConfig";
import Dashboard from "./Dashboard";
import WeeklyProductionReport from "./WPReport";
import ReportsMain from "./ReportsMain";

/* Development URL */
//  const webURL = "https://ggsaus.sharepoint.com/sites/Intranet_Test";
//  const WeblistURL = "Annual Plan Test";
//  const playbookURL =
//    "https://ggsaus.sharepoint.com/sites/Intranet_dev/SitePages/Playbook.aspx?activityID=";

/* Production URL */
// const webURL = "https://ggsaus.sharepoint.com";
// const WeblistURL = "Annual Plan";
// const playbookURL =
//   "https://ggsaus.sharepoint.com/sites/Intranet_Production/SitePages/Playbook.aspx?activityID=";

let webURL: string;
let WeblistURL: string;
let playbookURL: string;

if (window.location.href.toLowerCase().indexOf("production/") > -1) {
  /* Production URL */
  webURL = "https://ggsaus.sharepoint.com";
  WeblistURL = "Annual Plan";
  playbookURL =
    "https://ggsaus.sharepoint.com/sites/Intranet_Production/SitePages/Playbook.aspx?activityID=";
} else {
  /* Development URL */
  webURL = "https://ggsaus.sharepoint.com/sites/Intranet_Test";
  WeblistURL = "Annual Plan Test";
  playbookURL =
    "https://ggsaus.sharepoint.com/sites/Intranet_dev/SitePages/Playbook.aspx?activityID=";
}

let allPeoples = [];
const MainComponent = (props: any) => {
  const _webURL = Web(webURL);

  const [pageSwitch, setPageSwitch] = useState("");
  const [APID, setAPID] = useState();
  const [adminStatus, setAdminStatus] = useState(null);
  const [adminOrg, setAdminOrg] = useState(null);
  const [adminStock, setAdminStock] = useState(null);
  const [pageNavType, setpageNavType] = useState();
  const [pbSwitch, setpbSwitch] = useState();
  const [siteUsers, setSiteUsers] = useState(allPeoples);

  //let PageName;
  const handleclick = (
    page: string,
    AP_ID: any,
    navType: any,
    pbswitch: any
  ) => {
    // PageName = ActivityProductionBoard;
    setPageSwitch(page);
    setAPID(AP_ID);
    setpageNavType(navType);
    setpbSwitch(pbswitch);
  };
  const getAllUsers = () => {
    _webURL.siteUsers().then((_allUsers) => {
      _allUsers.forEach((user) => {
        let userName = user.Title.toLowerCase();
        // if (userName.indexOf("archive") == -1) {
        allPeoples.push({
          key: 1,
          imageUrl:
            `/_layouts/15/userphoto.aspx?size=S&accountname=` + `${user.Email}`,
          text: user.Title,
          ID: user.Id,
          secondaryText: user.Email,
          isValid: true,
        });
        // }
      });
      setSiteUsers([...allPeoples]);
      getAdmins();
    });
  };
  const getAdmins = () => {
    _webURL.siteGroups
      .getByName("Innovation Hub Admin")
      .users.get()
      .then((users) => {
        props.context.web
          .currentUser()
          .then((user) => {
            let tempUser = users.filter((_user) => {
              return (
                _user.Email == user.Email ||
                _user.Title == "Everyone except external users"
              );
            });
            if (tempUser.length > 0) {
              setAdminStatus(true);
            } else {
              setAdminStatus(false);
            }
            getOrgAdmins();
          })
          .catch((error) => {
            alert(error);
          });
      })
      .catch((error) => {
        alert(error);
      });
  };
  const getOrgAdmins = () => {
    _webURL.siteGroups
      .getByName("Org Report Admin")
      .users.get()
      .then((users) => {
        props.context.web
          .currentUser()
          .then((user) => {
            let tempUser = users.filter((_user) => {
              return (
                _user.Email == user.Email ||
                _user.Title == "Everyone except external users"
              );
            });
            if (tempUser.length > 0) {
              setAdminOrg(true);
            } else {
              setAdminOrg(false);
            }
            getStockAdmins();
          })
          .catch((error) => {
            alert(error);
          });
      })
      .catch((error) => {
        alert(error);
      });
  };
  const getStockAdmins = () => {
    _webURL.siteGroups
      .getByName("Stock List Admin")
      .users.get()
      .then((users) => {
        props.context.web
          .currentUser()
          .then((user) => {
            let tempUser = users.filter((_user) => {
              return (
                _user.Email == user.Email ||
                _user.Title == "Everyone except external users"
              );
            });
            if (tempUser.length > 0) {
              setAdminStock(true);
            } else {
              setAdminStock(false);
            }
            getDRAdmins();
          })
          .catch((error) => {
            alert(error);
          });
      })
      .catch((error) => {
        alert(error);
      });
  };
  const getDRAdmins = () => {
    _webURL.siteGroups
      .getByName("Document Review Admins")
      .users.get()
      .then((users) => {
        props.context.web
          .currentUser()
          .then((user) => {
            let tempUser = users.filter((_user) => {
              return (
                _user.Email == user.Email ||
                _user.Title == "Everyone except external users"
              );
            });
            if (tempUser.length > 0) {
              setPageSwitch("DocumentReview");
            } else {
              setPageSwitch("AnnualPlan");
            }
            pageFunction();
          })
          .catch((error) => {
            alert(error);
          });
      })
      .catch((error) => {
        alert(error);
      });
  };
  const pageFunction = () => {
    const urlParams = new URLSearchParams(window.location.search);
    const pageName = urlParams.get("Page");

    if (pageName == "AP") {
      setPageSwitch("AnnualPlan");
    } else if (pageName == "PB") {
      setPageSwitch("ProductionBoard");
    } else if (pageName == "DR") {
      setPageSwitch("DocumentReview");
    } else if (pageName == "ATP") {
      setPageSwitch("ActivityPlan");
    } else if (pageName == "PL") {
      setPageSwitch("PL_MainComponent");
    } else if (pageName == "MP") {
      setPageSwitch("MasterProduct");
    } else if (pageName == "OR") {
      setPageSwitch("OrgReporting");
    } else if (pageName == "SOL") {
      setPageSwitch("StockList");
    } else if (pageName == "SL") {
      setPageSwitch("StaffList");
    } 
    else if (pageName == "REP") {
      setPageSwitch("AllReports");
    }else if (pageName == "DL") {
      setPageSwitch("DistributionList");
    } else if (pageName == "DAC") {
      setPageSwitch("DistributionApprovalConfig");
    }else if (pageName == "DB") {
      setPageSwitch("Dashboard");
    }else if (pageName == "WPR") {
      setPageSwitch("WeeklyProductionReport");
    }
  };
  useEffect(() => {
    getAllUsers();
  }, []);
  return (
    <>
      {pageSwitch != "" && (adminStock == true || adminStock == false) ? (
        pageSwitch == "AnnualPlan" ? (
          <AnnualPlan
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            WeblistURL={WeblistURL}
            isAdmin={adminStatus}
            handleclick={handleclick}
            peopleList={siteUsers}
          />
        ) : pageSwitch == "DeliveryPlan" ? (
          <DeliveryPlan
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            WeblistURL={WeblistURL}
            handleclick={handleclick}
            AnnualPlanId={APID}
            peopleList={siteUsers}
            playbookURL={playbookURL}
          />
        ): pageSwitch == "AllReports" ? 
        (<ReportsMain 
          context={props.context}
              spcontext={props.spcontext}
              graphContent={props.graphContent}
              URL={webURL}
              handleclick={handleclick}
              pageType={pageNavType}
              peopleList={siteUsers}
              isAdmin={adminOrg}
          
          />) : pageSwitch == "ProductionBoard" ? (
          <ProductionBoard
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            WeblistURL={WeblistURL}
            handleclick={handleclick}
            AnnualPlanId={APID}
            pageType={pageNavType}
            pbSwitch={pbSwitch}
            peopleList={siteUsers}
          />
        ) : pageSwitch == "DocumentReview" ? (
          <DocumentReview
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            handleclick={handleclick}
            peopleList={siteUsers}
            isAdmin={adminStatus}
            // distributionListHandlerFunction={distributionListHandler}
          />
        ) : pageSwitch == "WeeklyProductionReport" ? (
          <WeeklyProductionReport
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            handleclick={handleclick}
            peopleList={siteUsers}
            isAdmin={adminStatus}
            // distributionListHandlerFunction={distributionListHandler}
          />
        ) : pageSwitch == "Dashboard" ? (
          <Dashboard
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            AnnualPlanId={APID}
            WeblistURL={WeblistURL}
            isAdmin={adminStatus}
            handleclick={handleclick}
            peopleList={siteUsers}
          />
        ) : pageSwitch == "ActivityPlan" ? (
          <ActivityPlan
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            handleclick={handleclick}
            peopleList={siteUsers}
            isAdmin={adminStatus}
            AnnualPlanId={APID}
            pbSwitch={pbSwitch}
            WeblistURL={WeblistURL}
          />
        ) : pageSwitch == "ActivityDeliveryPlan" ? (
          <ActivityDeliveryPlan
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            ActivityPlanID={APID}
            handleclick={handleclick}
            peopleList={siteUsers}
          />
        ) : pageSwitch == "ActivityTemplate" ? (
          <ActivityTemplate
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            WeblistURL={WeblistURL}
            handleclick={handleclick}
            peopleList={siteUsers}
          />
        ) : pageSwitch == "ActivityProductionBoard" ? (
          <ActivityProductionBoard
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            WeblistURL={WeblistURL}
            handleclick={handleclick}
            ActivityPlanID={APID}
            pageType={pageNavType}
            pbSwitch={pbSwitch}
            peopleList={siteUsers}
          />
        ) : pageSwitch == "PL_MainComponent" ? (
          <PL_MainComponent
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            WeblistURL={WeblistURL}
            handleclick={handleclick}
            peopleList={siteUsers}
          />
        ) : // ) : pageSwitch == "ProductActivityTemplate" ? (
        //   <ProductActivityTemplate
        //     context={props.context}
        //     spcontext={props.spcontext}
        //     graphContent={props.graphContent}
        //     URL={webURL}
        //     handleclick={handleclick}
        //     peopleList={siteUsers}
        //   />
        // ) : pageSwitch == "ProductActivityPlan" ? (
        //   <ProductActivityPlan
        //     context={props.context}
        //     spcontext={props.spcontext}
        //     graphContent={props.graphContent}
        //     URL={webURL}
        //     handleclick={handleclick}
        //     ActivityPlanID={APID}
        //     peopleList={siteUsers}
        //   />
        // ) : pageSwitch == "ProductActivityDeliveryPlan" ? (
        //   <ProductActivityDeliveryPlan
        //     context={props.context}
        //     spcontext={props.spcontext}
        //     graphContent={props.graphContent}
        //     URL={webURL}
        //     handleclick={handleclick}
        //     ActivityPlanID={APID}
        //     pageType={pageNavType}
        //     peopleList={siteUsers}
        //   />
        pageSwitch == "MasterProduct" ? (
          <MasterProduct
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            handleclick={handleclick}
            pageType={pageNavType}
            peopleList={siteUsers}
          />
        ) : pageSwitch == "DocumentReviewHistory" ? (
          <DocumentReviewHistory
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            handleclick={handleclick}
            pageType={pageNavType}
            peopleList={siteUsers}
            DRID={APID}
          />
        ) : pageSwitch == "OrgReporting" ? (
          <OrgReporting
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            handleclick={handleclick}
            pageType={pageNavType}
            peopleList={siteUsers}
            isAdmin={adminOrg}
          />
        ) : pageSwitch == "StockList" ? (
          <StockList
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            handleclick={handleclick}
            pageType={pageNavType}
            peopleList={siteUsers}
            isAdmin={adminStock}
            WeblistURL={WeblistURL}
          />
        ) : pageSwitch == "StaffList" ? (
          <StaffList
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            handleclick={handleclick}
            pageType={pageNavType}
            peopleList={siteUsers}
            isAdmin={adminStatus}
          />
        ) : pageSwitch == "DistributionApprovalConfig" ? (
          <DistributionApprovalConfig
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={webURL}
            handleclick={handleclick}
            pageType={pageNavType}
            peopleList={siteUsers}
            isAdmin={adminStatus}
          />
        ) : null
      ) : null}
    </>
  );
};

export default MainComponent;
