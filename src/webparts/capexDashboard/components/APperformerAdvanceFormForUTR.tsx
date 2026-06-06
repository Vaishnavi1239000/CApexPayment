import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { useEffect, useState } from "react";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
//import 'bootstrap/dist/css/bootstrap.min.css';
import logo from "../assets/sona-comstarlogo.png";
interface IProps {
  context: any;
  itemId: number; 
}
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}
const APperformerAdvanceFormForUTR: React.FC<IProps> = ({
  context,
  itemId,
}) => {const actionLock = React.useRef(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const sp = spfi().using(SPFx(context));
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [attachments, setAttachments] = useState<any[]>([]);
  const today = new Date();
const localDate: string = new Date(
  today.getTime() - today.getTimezoneOffset() * 60000
)
  .toISOString()
  .split("T")[0];
  const [employee, setEmployee] = useState<any>({});
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [UTRRemarks, setUTRRemarks] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  // ✅ Fetch Item by ID
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
const getLoggedInUser = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      const email = currentUser.Email;

      const user = await sp.web.lists
        .getByTitle("EmployeeMaster")
        .items.select(
          "EmployeeCode",
          "EmployeeName",
          "Division",
          "Location",
          "EmployeeEmail",
          "ReportingManager/Title",
          "HOD/Title",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();

      if (user.length > 0) {
        setEmployee(user[0]);
      }
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };
  const getPreviousAdvances = async (vendorId: number) => {
    try {
      debugger;
       if (!vendorId) {
      setPreviousAdvances([]);
      return;
    }
      console.log("Fetching for Vendor:", vendorId);

      const data = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "PONumber",
          "RequestAdvanceAmount",
          "Created",
          "VoucherDate",

          "PaidAmount",
          "Status",
          "VendorCode/Id",
        )
        .expand("VendorCode")
        .filter(`VendorCode/Id eq ${vendorId} and Status eq 'Paid'`)
        .orderBy("Created", false)();

      console.log("DATA:", data);

      void setPreviousAdvances(data);
    } catch (error) {
      console.error("Error fetching previous advances:", error);
      void setPreviousAdvances([]);
    }
  };
  const getAttachments = async (capexId: string) => {
    try {
      const safe = capexId.replace(/\//g, "_");
      const path = `/sites/SonaFinance/CapexPaymentDocs/${safe}`;

      const files = await sp.web.getFolderByServerRelativePath(path).files();

      void setAttachments(files);
    } catch {
      void setAttachments([]);
    }
  };
  const getVendors = async () => {
    try {
      const data = await sp.web.lists
        .getByTitle("VendorMaster")
        .items.select("Id", "VendorCode", "VendorName")();

      setVendors(data);
    } catch (error) {
      console.error("Vendor fetch error:", error);
    }
  };
  // ✅ Fetch Item by ID
  
  const getItemById = async (itemId: any) => {
    try {
      debugger;
      console.log("Fetching item with ID:", itemId);

      const item = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemId)
        .select(
          "ID",
          "Title",
          "Created",
          "EmployeeName",
          "Email",
          "EmployeeCode",
          "ContactNo",
          "EmployeeStatus",
          "Division",
          "Location",
          "RM",
          "HOD",
          "VendorCode/Id",
          "VendorCode/Title",
          "VendorName",
          "PODate",
          "POPaymentTerms",
          "POAmount",
          "MRNNumber",
          "MRNDtae",
          "MRNAmountwithGST",
          "PONumber",
          "RequestedAmountforPayment",
          "Status",
          "CurrentApproverId",
          "CapexId",
          "VoucherDate",
          "VoucherNumber",
          "ApproverRemarks",
           "InstallationDetails",
          "FinalPaymentAgainstPO",
          "ApprovalMatrix",
          "WorkflowHistory",
          
         "RequestorName"

        )();

      console.log("ITEM DATA:", item.RequestorName);
      debugger;
      setItemData(item);
      debugger;

      const vendorId = item?.VendorCode?.Id || null;

      console.log("Vendor Id:", vendorId);

      setSelectedVendorId(vendorId);

      setSelectedVendorName(item?.VendorName || "");

      if (item.CapexId) {
        await getAttachments(item.CapexId);
      }

      if (item.ApprovalMatrix) {
        try {
          const parsed =
            typeof item.ApprovalMatrix === "string"
              ? JSON.parse(item.ApprovalMatrix)
              : item.ApprovalMatrix;

          setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
        } catch (e) {
          console.error("ApprovalMatrix parse error", e);
          setApprovalMatrix([]);
        }
      } else {
        setApprovalMatrix([]);
      }

      if (item.WorkflowHistory) {
        try {
          const parsed =
            typeof item.WorkflowHistory === "string"
              ? JSON.parse(item.WorkflowHistory)
              : item.WorkflowHistory;

          setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
        } catch (e) {
          console.error("WorkFlowHistory parse error", e);
          setWorkflowHistory([]);
        }
      } else {
        setWorkflowHistory([]);
      }
    } catch (error) {
      console.error("Fetch error:", error);
    }
  };

useEffect(() => {
    if (!context || !itemId) return;
    debugger;
    const loadData = async () => {
      debugger;

      await getLoggedInUser();

      
      await getVendors();

     
      await getItemById(itemId);
    };

    void loadData();
  }, [context, itemId]);

  useEffect(() => {
    if (selectedVendorId) {
      console.log("Calling Previous Advances:", selectedVendorId);

      void getPreviousAdvances(selectedVendorId);
    }
  }, [selectedVendorId]);

  // ✅ Approve
  const handleApprove = async () => {

     if (actionLock.current) return;

    // actionLock.current = true;
    if (isSubmitting) return;
    setIsSubmitting(true);
    try {
       if (!UTRDate || UTRDate.trim() === "") {
        alert("Please enter UTR Date");
       setIsSubmitting(false);
        return;
      }
       if (UTRDate > localDate) {
      alert("UTR date cannot be a future date");
      // return;
       setIsSubmitting(false);
        return;
    }

      if (!UTRNumber || UTRNumber.trim() === "") {
        alert("Please enter UTR Number");
        setIsSubmitting(false);
        return;
      }

      if (!UTRRemarks || UTRRemarks.trim() === "") {
        alert("Please enter UTR Remarks");
       setIsSubmitting(false);
        return;
      }


      // 🔥 HISTORY
      const history = itemData.WorkflowHistory
        ? (typeof itemData.WorkflowHistory === "string"
          ? JSON.parse(itemData.WorkflowHistory)
          : itemData.WorkflowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Paid",
        Comment: UTRRemarks,
        Date: new Date().toISOString()
      });

      // 🔥 MATRIX
     
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,
          UTRDate: UTRDate ? new Date(UTRDate) : null,
          UTRNumber: UTRNumber,
          UTRRemarks: UTRRemarks,

          Status: "Paid",

         WorkflowHistory: JSON.stringify(history),
          //ApprovalMatrix: JSON.stringify(flow),

         // CurrentApproverId: null
        });

      alert("Paid successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Performer";

    } catch (error) {
      console.error("Approve error:", error);
      alert("Error ❌");
    }
    finally {
      setIsSubmitting(false);
    }
  };

  // ✅ Sent Back
  const handleSendBack = async () => {
     if (actionLock.current) return;
    // actionLock.current = true;
    if (isSubmitting) return;
    try {
     if (!UTRRemarks || UTRRemarks.trim() === "") {
        alert("Please enter UTR Remarks");
       setIsSubmitting(false);
        return;
      }

      const history = itemData.WorkflowHistory
        ? (typeof itemData.WorkflowHistory === "string"
          ? JSON.parse(itemData.WorkflowHistory)
          : itemData.WorkflowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Send Back",
        Comment: UTRRemarks,
        Date: new Date().toISOString()
      });

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex(
        (a: any) => a.Id === currentUserId
      );

      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Send Back";
      }

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Send Back",

         WorkflowHistory: JSON.stringify(history),
          ////ApprovalMatrix: JSON.stringify(flow),

         // CurrentApproverId: itemData.CurrentApproverId
        });

      alert("Send Back ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Performer";

    } catch (error) {
      console.error(error);
    }
    finally {
      setIsSubmitting(false);
    }
  };

  // ✅ Reject
  const handleReject = async () => {
     if (actionLock.current) return;
    // actionLock.current = true;
    if (isSubmitting) return;
    try {
      if (!UTRRemarks || UTRRemarks.trim() === "") {
        alert("Please enter UTR Remarks");
        setIsSubmitting(false);
        return;
      }

      const history = itemData.WorkflowHistory
        ? (typeof itemData.WorkflowHistory === "string"
          ? JSON.parse(itemData.WorkflowHistory)
          : itemData.WorkflowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Rejected",
        Comment: UTRRemarks,
        Date: new Date().toISOString()
      });

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex(
        (a: any) => a.Id === currentUserId
      );

      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Rejected";
      }

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Rejected",

         WorkflowHistory: JSON.stringify(history),
         // ApprovalMatrix: JSON.stringify(flow),

       //   CurrentApproverId: itemData.CurrentApproverId
        });

      alert("Rejected ❌");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Performer";

    } catch (error) {
      console.error(error);
    }
      finally { 
      setIsSubmitting(false);
  }
  };

  const handleExit = () => {
    window.location.href = `https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Performer`;
  };

  // ⛔ Wait until data loads
  if (!itemData) return <div>Loading...</div>;

  return (
    <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
      <div className='row'>
        <div className='col-md-12'>
          <div className='Main-Boxpoup'>
            <div className="bordered">
              <img src={logo} />
              <h1> Advance Payment (Approver) </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <p>No approval data</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  <li className={`approval-step`}>
                    {`Initiator`} - {itemData?.EmployeeName}
                  </li>
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${
                        a.Status === "In Progress"
                          ? "active"
                          : a.Status === "Approved"
                            ? "approved"
                            : a.Status === "Rejected"
                              ? "rejected"
                              : a.Status === "Send Back"
                                ? "sendback"
                                : ""
                      }`}
                    >
                      {a.Role} - {a.Name}
                    </li>
                  ))}
                </ul>
              </div>
            )}
            <div className='borderedbox'>
              <div className="heading1">
                <label>Requestor Information</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Code" className='font'>Employee Code</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeCode}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Name" className='font'>Employee Name </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeName}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Email" className='font'>Employee Email </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.Email}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Contact No" className='font'>Contact No</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.ContactNo}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Status" className='font'>Employee Status</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeStatus}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Division" className='font'>Division</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.Division}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Location" className='font'>Location</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.Location}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="RM" className='font'>RM</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.RM}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="HOD" className='font'>HOD</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.HOD}</label>
                  </div>
                </div>
              </div>
              <div className="heading1">
                <label>Vendor & PO Details</label>
              </div>
               <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font"> Vendor Code </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.VendorCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.VendorName}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.PONumber}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.PODate ? new Date(itemData.PODate).toLocaleDateString("en-GB",) : ""}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Terms </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.POPaymentTerms}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.POAmount}</label>
                  </div>
                </div>
              </div>
              -
              <div className="heading1">
                <label>Approver Action</label>
              </div>
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Approver Remarks</label>
                     <input value={itemData.ApproverRemarks || ""} className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Voucher Date</label>
                    <input value={itemData.VoucherDate
                      ? new Date(itemData.VoucherDate).toLocaleDateString("en-GB") : ""}
                      className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Voucher Number</label>
                    <input value={itemData.VoucherNumber || ""} className="font-control readonly" />
                  </div>
                </div>
              </div>

                <div className="heading1" style={{ marginTop: "10px" }}>
                <label>MRN & Payment Details </label> : &nbsp;&nbsp;
              </div>

              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Number </label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData?.MRNNumber}</label>
                  </div>

                  <div className="col-md-4">
                    <label className="font">MRN Date</label> : &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.MRNDtae
                        ? new Date(itemData.MRNDtae).toLocaleDateString("en-GB")
                        : ""}
                    </label>
                    {/* <label className="fonttext">{itemData?.mrnDate}</label> */}
                  </div>

                  <div className="col-md-4">
                    <label className="font">MRN Amount</label> : &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.MRNAmountwithGST}
                    </label>
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Requested Amount</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.RequestedAmountforPayment}
                    </label>
                  </div>
                </div>
              </div>
              <div className="heading1">
                <label>Upload Document</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Attachments</label>
                    {attachments.length === 0 ? (
                      <p>No attachments</p>
                    ) : (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a
                              href={file.ServerRelativeUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {file.Name}
                            </a>
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                  <div className="col-md-4">
                    <label className="font">UTR Date</label>
                    <input type="date" className="font-control" value={UTRDate}
                      max={new Date().toISOString().split("T")[0]}
                      onChange={(e) => setUTRDate(e.target.value)} />
                  </div>
                  <div className="col-md-4">
                    <label className="font">UTR Number</label>
                    <input value={UTRNumber} className="font-control" onChange={(e) => setUTRNumber(e.target.value)} />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">UTR Remarks</label>
                    <input className="font-control" onChange={(e) => setUTRRemarks(e.target.value)} />
                  </div>
                </div>
              </div>
               <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Workflow History</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    {workflowHistory.length === 0 ? (
                      <p>No history available</p>
                    ) : (
                      <div className="workflow-history">
                        

                        <table
                          className="workflow-table"
                          style={{ width: "100%" }}
                        >
                          <thead>
                            <tr>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Action By
                              </th>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Action Taken
                              </th>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Date
                              </th>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Comment
                              </th>
                            </tr>
                          </thead>
                          <tbody>
                            {workflowHistory
                              .filter(
                                (h: any) =>
                                  h.ActionTaken &&
                                  h.ActionTaken !== "Draft Saved" && h.ActionTaken !== "Edited",
                              )
                              .map((h: any, idx: number) => (
                                <tr key={idx}>
                                  <td style={{ padding: "8px" }}>
                                    {h.CurrentApprover || ""}
                                  </td>

                                  <td style={{ padding: "8px" }}>
                                    {h.ActionTaken || ""}
                                  </td>

                                  <td style={{ padding: "8px" }}>
                                    {h.Date
                                      ? new Date(h.Date).toLocaleDateString(
                                          "en-GB",
                                        )
                                      : ""}
                                  </td>

                                  <td style={{ padding: "8px" }}>
                                    {h.Comment || ""}
                                  </td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                       
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div className='row my-3'>
                <div className='col-md-12'>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "5px" }}>
                   <a
                          className={`submit-btn ${isSubmitting ? "disabled-btn" : ""}`}
                          onClick={!isSubmitting ? handleApprove : undefined}
                        >
                          {isSubmitting ? "Processing..." : "Paid"}
                        </a>

                        <a
                          className={`Rework-btn ${isSubmitting ? "disabled-btn" : ""}`}
                          onClick={!isSubmitting ? handleSendBack : undefined}
                        >
                          {isSubmitting ? "Processing..." : "Send Back"}
                        </a>

                        <a
                          className={`Reject-btn ${isSubmitting ? "disabled-btn" : ""}`}
                          onClick={!isSubmitting ? handleReject : undefined}
                        >
                          {isSubmitting ? "Processing..." : "Reject"}
                        </a>
                    <a href="#" onClick={handleExit} className="reset-btn">
                      Exit
                    </a>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div >
  );
};

export default APperformerAdvanceFormForUTR;
