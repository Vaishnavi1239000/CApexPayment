import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
//import "bootstrap/dist/css/bootstrap.min.css";
import logo from "../assets/sona-comstarlogo.png";
interface IProps {
  context: any;
  formData: any;        // ✅ ADD THIS
  onClose: () => void;  // ✅ ADD THIS (if not present)
}
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}
const APperformerAdvanceform: React.FC<IProps> = ({ context, formData, onClose }) => {

  const sp = spfi().using(SPFx(context));

  const [employee, setEmployee] = useState<any>({});
  const [attachments, setAttachments] = useState<any[]>([]);
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [voucherNumber, setVoucherNumber] = useState("");
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
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
  // ✅ Fetch Item by ID

  useEffect(() => {
    if (!formData) return;

    setItemData(formData);
    setApproverRemarks(formData.ApproverRemarks || "");

    // ✅ Workflow History
    if (formData.WorkflowHistory) {
      try {
        setWorkflowHistory(formData.WorkflowHistory);
      } catch {
        setWorkflowHistory([]);
      }
    }
    if (formData.ApprovalMatrix) {
      try {
        setApprovalMatrix(formData.ApprovalMatrix);
      } catch {
        setApprovalMatrix([]);
      }
    }

    // 🔥 ADD THIS LINE (MAIN FIX)
    if (formData.Title) {
      getAttachments(formData.Title);
    }

  }, [formData]);

  // ✅ Approve
  const handleApprove = async () => {

    try {

      if (!voucherDate) {
        alert("Please select Voucher Date");
        return;
      }

      if (!voucherNumber) {
        alert("Please enter Voucher Number");
        return;
      }

      if (!approverRemarks) {
        alert("Please enter Remarks");
        return;
      }

      // =========================
      // 🔥 PARSE APPROVAL MATRIX
      // =========================
      let flow: any[] = [];

      try {

        flow =
          typeof itemData.ApprovalMatrix === "string"
            ? JSON.parse(itemData.ApprovalMatrix)
            : itemData.ApprovalMatrix || [];

      } catch {

        flow = [];
      }

      if (!Array.isArray(flow) || flow.length === 0) {
        alert("Approval Matrix empty");
        return;
      }

      // =========================
      // 🔥 FIND CURRENT APPROVER
      // =========================
      const currentIndex = flow.findIndex(
        (x: any) =>
          x.Status === "Pending" ||
          x.Status === "In Progress"
      );

      if (currentIndex === -1) {
        alert("No pending approver");
        return;
      }

      // =========================
      // 🔥 CURRENT APPROVED
      // =========================
      flow[currentIndex].Status = "Approved";

      // =========================
      // 🔥 NEXT APPROVER
      // =========================
      const nextIndex = currentIndex + 1;

      let nextApproverId: number | null = null;

      let nextStatus = "Approved";

      let pendingAt = "Completed";

      if (nextIndex < flow.length) {

        flow[nextIndex].Status = "Pending";

        nextApproverId = flow[nextIndex].Id;

        const nextRole = flow[nextIndex].Role;



        nextStatus = "Pending for PF Approver UTR";



        pendingAt = `Pending at ${nextRole}`;
      }

      // =========================
      // 🔥 WORKFLOW HISTORY
      // =========================
      let history: any[] = [];

      try {

        history =
          typeof itemData.WorkflowHistory === "string"
            ? JSON.parse(itemData.WorkflowHistory)
            : itemData.WorkflowHistory || [];

      } catch {

        history = [];
      }

      history.push({

        CurrentApprover:
          context.pageContext.user.displayName,

        ActionTaken: "Vouched",

        Comment: approverRemarks,

        Date: new Date().toISOString(),

        CurrentStatus: pendingAt

      });

      console.log("🔥 UPDATED FLOW:", flow);
      console.log("🔥 UPDATED HISTORY:", history);

      // =========================
      // 🔥 UPDATE SHAREPOINT
      // =========================
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({

          ApproverRemarks: approverRemarks,

          VoucherDate:
            voucherDate
              ? new Date(voucherDate)
              : null,

          VoucherNumber: voucherNumber,

          Status: nextStatus,

          PendingAt: pendingAt,

          // 🔥 IMPORTANT
          ApprovalMatrix: JSON.stringify(flow),

          // 🔥 IMPORTANT
          WorkflowHistory: JSON.stringify(history),

          // 🔥 IMPORTANT
          // CurrentApproverIdId: nextApproverId
        });

      alert("Vouched successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Performer";

    } catch (error) {

      console.error("Approve error:", error);

      alert("Error ❌");
    }
  };

  // ✅ Sent Back
  const handleSendBack = async () => {

    try {

      if (!approverRemarks) {
        alert("Please enter Remarks");
        return;
      }

      // =========================
      // 🔥 PARSE MATRIX
      // =========================
      let flow: any[] = [];

      try {

        flow =
          typeof itemData.ApprovalMatrix === "string"
            ? JSON.parse(itemData.ApprovalMatrix)
            : itemData.ApprovalMatrix || [];

      } catch {

        flow = [];
      }

      if (!Array.isArray(flow) || flow.length === 0) {
        alert("Approval Matrix empty");
        return;
      }

      // =========================
      // 🔥 FIND CURRENT
      // =========================
      const currentIndex = flow.findIndex(
        (x: any) =>
          x.Status === "Pending" ||
          x.Status === "In Progress"
      );

      if (currentIndex === -1) {
        alert("No current approver found");
        return;
      }

      // =========================
      // 🔥 SEND BACK
      // =========================
      flow[currentIndex].Status = "Send Back";

      let previousApproverId: number | null = null;

      if (currentIndex > 0) {

        flow[currentIndex - 1].Status = "Pending";

        previousApproverId = flow[currentIndex - 1].Id;
      }

      // =========================
      // 🔥 WORKFLOW HISTORY
      // =========================
      let history: any[] = [];

      try {

        history =
          typeof itemData.WorkflowHistory === "string"
            ? JSON.parse(itemData.WorkflowHistory)
            : itemData.WorkflowHistory || [];

      } catch {

        history = [];
      }

      history.push({

        CurrentApprover:
          context.pageContext.user.displayName,

        ActionTaken: "Send Back",

        Comment: approverRemarks,

        Date: new Date().toISOString(),

        CurrentStatus: "Send Back"
      });

      console.log("🔥 UPDATED FLOW:", flow);
      console.log("🔥 UPDATED HISTORY:", history);

      // =========================
      // 🔥 UPDATE SHAREPOINT
      // =========================
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({

          ApproverRemarks: approverRemarks,

          Status: "Send Back",

          PendingAt:
            currentIndex > 0
              ? `Pending at ${flow[currentIndex - 1].Role}`
              : "Send Back",

          // ApprovalMatrix: JSON.stringify(flow),

          WorkflowHistory: JSON.stringify(history),

          // CurrentApproverIdId: previousApproverId
        });

      alert("Send Back ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Performer";

    } catch (error) {

      console.error(error);

      alert("Error ❌");
    }
  };

  // ✅ Reject
  const handleReject = async () => {

    try {

      if (!approverRemarks) {
        alert("Please enter Remarks");
        return;
      }

      // =========================
      // 🔥 PARSE MATRIX
      // =========================
      let flow: any[] = [];

      try {

        flow =
          typeof itemData.ApprovalMatrix === "string"
            ? JSON.parse(itemData.ApprovalMatrix)
            : itemData.ApprovalMatrix || [];

      } catch {

        flow = [];
      }

      if (!Array.isArray(flow) || flow.length === 0) {
        alert("Approval Matrix empty");
        return;
      }

      // =========================
      // 🔥 FIND CURRENT
      // =========================
      const currentIndex = flow.findIndex(
        (x: any) =>
          x.Status === "Pending" ||
          x.Status === "In Progress"
      );

      if (currentIndex === -1) {
        alert("No current approver found");
        return;
      }

      // =========================
      // 🔥 REJECTED
      // =========================
      flow[currentIndex].Status = "Rejected";

      // =========================
      // 🔥 WORKFLOW HISTORY
      // =========================
      let history: any[] = [];

      try {

        history =
          typeof itemData.WorkflowHistory === "string"
            ? JSON.parse(itemData.WorkflowHistory)
            : itemData.WorkflowHistory || [];

      } catch {

        history = [];
      }

      history.push({

        CurrentApprover:
          context.pageContext.user.displayName,

        ActionTaken: "Rejected",

        Comment: approverRemarks,

        Date: new Date().toISOString(),

        CurrentStatus: "Rejected"
      });

      console.log("🔥 UPDATED FLOW:", flow);
      console.log("🔥 UPDATED HISTORY:", history);

      // =========================
      // 🔥 UPDATE SHAREPOINT
      // =========================
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({

          ApproverRemarks: approverRemarks,

          Status: "Rejected",

          PendingAt: "Rejected",

          // ApprovalMatrix: JSON.stringify(flow),

          WorkflowHistory: JSON.stringify(history),

          // CurrentApproverIdId: null
        });

      alert("Rejected ❌");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Performer";

    } catch (error) {

      console.error(error);

      alert("Error ❌");
    }
  };
  const handleExit = () => {
    window.location.href = `https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Performer`;
  };

  // ⛔ Wait until data loads
  if (!itemData) return <div>Loading...</div>;

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            <div className="bordered">
              <img src={logo} />
              <h1> Advance Payment (Approver) </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <p></p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${a.Status === "In Progress"
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
            <div className="borderedbox">
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Employee Code" className="font">
                      Employee Code
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">
                      Employee Name{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">
                      Employee Email{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {itemData.Email}
                    </label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">
                      Contact No
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">
                      Employee Status
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {itemData.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">
                      Division
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">
                      Location
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">
                      RM
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {itemData.RM}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">
                      HOD
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.HOD}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
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
                    <label className="fonttext "> {itemData.POAdvanceTerms}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.POAmtGST}</label>
                  </div>
                </div>
              </div>
              <div className="main-formcontainer" style={{ marginTop: "10px" }}>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Voucher Date</label>
                    <input
                      type="date"
                      value={voucherDate}
                      onChange={(e) => setVoucherDate(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Voucher Number</label>
                    <input
                      value={voucherNumber}
                      onChange={(e) => setVoucherNumber(e.target.value)}
                      className="form-control"
                    />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Upload Document</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
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
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Workflow History</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    {/* {workflowHistory.length === 0 ? (
                      <p>No history available</p>
                    ) : (
                      <div className="workflow-history">
                        {workflowHistory.map((h, index) => (
                          <div key={index} className="history-item">
                            <div>
                              {h.ActionTaken === "Approved" && "✅ "}
                              {h.ActionTaken === "Rejected" && "❌ "}
                              {h.ActionTaken === "Send Back" && "↩ "}
                              {h.ActionTaken === "Vouched" && "💰 "}
                              {h.ActionTaken}
                            </div>

                            <div>
                              <b>{h.CurrentApprover}</b>
                            </div>
                            <div>{h.Comment}</div>
                            <div className="date">
                              {new Date(h.Date).toLocaleString()}
                            </div>
                          </div>
                        ))}
                      </div>
                    )} */}
                    <div className='Workflowbox'>
                      {workflowHistory && workflowHistory.length > 0 ? (
                        <table className="workflow-table" style={{ width: '100%', borderCollapse: 'collapse' }}>
                          <thead>
                            <tr>
                              <th style={{ padding: '8px', textAlign: 'left' }}>Action Date</th>
                              <th style={{ padding: '8px', textAlign: 'left' }}>Action By</th>
                              <th style={{ padding: '8px', textAlign: 'left' }}>Action Taken</th>
                              <th style={{ padding: '8px', textAlign: 'left' }}>Comment</th>
                            </tr>
                          </thead>
                          <tbody>
                            {workflowHistory.map((h: any, idx: number) => (
                              <tr key={idx}>
                                <td style={{ padding: '8px' }}>{h.Date ? new Date(h.Date).toLocaleString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }).replace(",", "") : ""}</td>
                                <td style={{ padding: '8px' }}>{h.CurrentApprover || ''}</td>
                                <td style={{ padding: '8px' }}>{h.ActionTaken || ''}</td>
                                <td style={{ padding: '8px' }}>{h.Comment || ''}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      ) : (
                        <p>No workflow history</p>
                      )}
                    </div>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Approver Action</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-6">
                    <label className="font">Approver Remarks</label>
                    <textarea
                      className="form-control "
                      value={approverRemarks}
                      onChange={(e) => setApproverRemarks(e.target.value)}
                    />
                  </div>
                </div>
              </div>
              <div className="row my-3">
                <div className="col-md-12">
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: "5px",
                    }}
                  >
                    <a className="submit-btn" onClick={handleApprove}>
                      Submit
                    </a>

                    <a className="Rework-btn" onClick={handleSendBack}>
                      send Back
                    </a>

                    <a className="Reject-btn" onClick={handleReject}>
                      Reject
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
    </div>
  );
};

export default APperformerAdvanceform;
