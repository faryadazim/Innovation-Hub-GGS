import * as React from 'react'


const MainComponents = () => {

  const [pgeSwitch, setPageSwitch] = React.useState("")
  const [businessAreaWise, setBusinessAreaWise]: any = React.useState<any>([])
  const [productionBoard, setProductionBoard] = React.useState([])
  const pageFunction = async () => {

    const urlParams = new URLSearchParams(window.location.search);
    const pageName = urlParams.get("Page");
    // const sharepointWeb = Web(webURL);
    let webURL = "https://ggsaus.sharepoint.com";
    const ap_list = "Annual Plan";
    const numberOfItems = 5000;
    const endpointUrl = `${webURL}/_api/web/lists/getbytitle('${ap_list}')/items?$top=${numberOfItems}`;

    const headers = new Headers();
    headers.append("Accept", "application/json;odata=verbose");

    fetch(endpointUrl, {
      method: "GET",
      headers: headers
    })
      .then(response => response.json())
      .then(res => {
        const data = res?.d?.results;

        const countsMap = new Map();

        data.forEach((item: any) => {
          const { BusinessArea, Status } = item;
          const lowercaseStatus = Status.toLowerCase();

          if (countsMap.has(BusinessArea)) {
            const statusCounts = countsMap.get(BusinessArea);
            statusCounts[lowercaseStatus] += 1;
          } else {
            countsMap.set(BusinessArea, {
              completed: lowercaseStatus === "completed" ? 1 : 0,
              behind_schedule: lowercaseStatus === "behind schedule" ? 1 : 0,
              scheduled: lowercaseStatus === "scheduled" ? 1 : 0,
            });
          }
        });
        let displayData: any = []
        countsMap.forEach((statusCounts, businessArea) => {
          const { completed, behind_schedule, scheduled } = statusCounts;
          displayData.push({
            businessArea,
            completed,
            behind_schedule,
            scheduled
          })

        });

        setBusinessAreaWise(displayData);


      })
      .catch(error => {
        // Handle any errors
        console.error("Error:", error);
      });


    const adp_list = "Activity Delivery Plan";
    // const endpointUrl2 = `${webURL}/_api/web/lists/getbytitle('${adp_list}')/$orderby=Modified desc&${numberOfItems}`;
  
    const endpointUrl2 = `${webURL}/_api/web/lists/getbytitle('${adp_list}')/items?$orderby=Modified desc&$top=${numberOfItems}`;
    fetch(endpointUrl2, {
      method: "GET",
      headers: headers
    })
      .then(response => response.json())
      .then(res => {
        const data = res?.d?.results;
 
    
        const currentDate = new Date();
        const currentWeek = getWeekNumber(currentDate);
 
        function getWeekNumber(date: any) {
          const d: any = new Date(date);
          d.setHours(0, 0, 0, 0);
          d.setDate(d.getDate() + 4 - (d.getDay() || 7));
          const yearStart: any = new Date(d.getFullYear(), 0, 1);
          const weekNo = Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
          return weekNo;
        }



        const filteredData = data.filter((obj: any) => {
          const StartDate = new Date(obj.StartDate);
          const EndDate = new Date(obj.EndDate);
 
          const startWeek = getWeekNumber(StartDate);
          const endWeek = getWeekNumber(EndDate);
          const isInCurrentWeek = currentWeek >= startWeek && currentWeek <= endWeek;
 
           const isPhZero = obj.ph === 0;
 
          return isInCurrentWeek && isPhZero ;
        });



        const uniqueArray: any = [];
        const uniqueNames: any = {};

        for (const item of filteredData) {
          if (item?.DeveloperId) {
            const DeveloperId = item.DeveloperId;
            if (!uniqueNames[DeveloperId]) {
              uniqueNames[DeveloperId] = true;
              uniqueArray.push(item);
            }
          }

        }


 
        setProductionBoard(uniqueArray)
      })
      .catch(error => {
        // Handle any errors
        console.error("Error:", error);
      });





    if (pageName == "INV") {
      setPageSwitch("INV");
    } else {
      setPageSwitch("PR");
    }
  };
  React.useEffect(() => {
    pageFunction()
  }, [])
  return (
    <>

      {
        pgeSwitch == "INV" ? <><>
          <h3>Innovation Hub</h3>
          <table style={{ borderCollapse: 'collapse', width: '100%' }}>
            <thead>
              <tr>
                <th style={{ padding: 8, textAlign: 'left', borderBottom: '1px solid #ddd', backgroundColor: '#f2f2f2' }}>Business Area</th>
                <th style={{ padding: 8, textAlign: 'left', borderBottom: '1px solid #ddd', backgroundColor: '#f2f2f2' }}>On Time</th>
                <th style={{ padding: 8, textAlign: 'left', borderBottom: '1px solid #ddd', backgroundColor: '#f2f2f2' }}>Overdue</th>
              </tr>
            </thead>
            <tbody>
              {
                businessAreaWise.map((rows: any) => {
                  return <tr>
                    <td style={{ padding: 8, textAlign: 'left', borderBottom: '1px solid #ddd' }}>{rows?.businessArea}</td>
                    <td style={{ padding: 8, textAlign: 'left', borderBottom: '1px solid #ddd' }}>{rows?.completed + rows?.scheduled}</td>
                    <td style={{ padding: 8, textAlign: 'left', borderBottom: '1px solid #ddd' }}>{rows?.behind_schedule}</td>
                  </tr>
                })
              }

            </tbody>
          </table>

        </></> : <>
          <h3>Production Board</h3>
          <table style={{ borderCollapse: 'collapse', width: '100%' }}>
            <thead>
              <tr>
                <th style={{ padding: 8, textAlign: 'left', borderBottom: '1px solid #ddd', backgroundColor: '#f2f2f2' }}>Name</th>
                 </tr>
            </thead>
            <tbody>
              {
                productionBoard?.map((rows: any) => {
                  return <tr>
                    <td style={{ padding: 8, textAlign: 'left', borderBottom: '1px solid #ddd' }}>Deveolper - id {rows?.DeveloperId}</td>
                   </tr>
                })
              }

            </tbody>
          </table></>
      }

    </>
  )
}

export default MainComponents