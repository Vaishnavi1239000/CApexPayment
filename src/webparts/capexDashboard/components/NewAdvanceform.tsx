import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState,useRef } from "react";
//import { Link, useHistory } from "react-router-dom";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import "bootstrap/dist/css/bootstrap.min.css";

import logo from "../assets/sona-comstarlogo.png";

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const NewAdvanceform = ({ context }: any) => {
  // const history = useHistory();
const submitRef = useRef(false);
  const draftRef = useRef(false);
  const sp = spfi().using(SPFx(context));
  const today = new Date();
  const localDate: string = new Date(
  today.getTime() - today.getTimezoneOffset() * 60000
)
  .toISOString()
  .split("T")[0];
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDraftSaving, setIsDraftSaving] = useState(false);
  const [employee, setEmployee] = React.useState<any>({});
  //const [selectedUser, setSelectedUser] = useState<any>(null);
  const [attachments, setAttachments] = useState<File[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);

  const [employeeName, setEmployeeName] = React.useState("");
  const [pickerKey, setPickerKey] = React.useState<number>(0);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [finalPayment, setFinalPayment] = useState("");
  const [installationDetails, setInstallationDetails] = useState("");

  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [poAmount, setPoAmount] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [paidAmount, setPaidAmount] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [mrnNumber, setMrnNumber] = useState("");
  const [mrnDate, setMrnDate] = useState("");
  const [mrnAmount, setMrnAmount] = useState("");
  const [requestedAmount, setRequestedAmount] = useState("");

  const [expectedDate, setExpectedDate] = useState("");

  const [glCode, setGlCode] = useState("390111001");

  const [costCenter, setCostCenter] = useState("");

  const [remarks, setRemarks] = useState("");
  const [Purpose, setPurpose] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const handleNumberChange = (value: string, setter: any) => {
    // Allow only numbers and decimal (max one dot)
    const regex = /^\d*\.?\d*$/;

    if (regex.test(value)) {
      void setter(value);
    }
  };
  const handleRemoveFile = (index: number) => {
    const updatedFiles = [...attachments];
    updatedFiles.splice(index, 1);
    setAttachments(updatedFiles);
  };

  const getPreviousAdvances = async (vendorId: number) => {
    try {
       if (!vendorId) {
      setPreviousAdvances([]);
      return;
    }
      debugger;
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

  const handleExit = () => {
    window.location.href =
      "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=User";
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
          "ReportingManager/Id",
          "HOD/Title",
          "HOD/Id",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();

      if (user.length > 0) {
        void setEmployee(user[0]);
      }
      buildApprovalPreview(user[0]);
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };
  const getVendors = async () => {
  try {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName", "Status")
      .filter("Status eq 'Active'")()
;

    setVendors(data);
  } catch (error) {
    console.error("Vendor fetch error:", error);
  }
};

  const getFinancialYear = () => {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;

    if (month >= 4) {
      // April to March
      return `${year.toString().slice(-2)}-${(year + 1).toString().slice(-2)}`;
    } else {
      return `${(year - 1).toString().slice(-2)}-${year.toString().slice(-2)}`;
    }
  };
  const generateCapexId = async () => {
    try {
      const fy = getFinancialYear();

      const items = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select("CapexId", "ID")
        .filter(`startswith(CapexId,'CPX-PYMT/${fy}/')`)
        .orderBy("ID", false)
        .top(1)();

      let nextNumber = 1;

      if (items.length > 0 && items[0].CapexId) {
        const lastId = items[0].CapexId;

        const parts = lastId.split("/");

        if (parts.length === 3) {
          const lastNumber = parseInt(parts[2]);

          if (!isNaN(lastNumber)) {
            nextNumber = lastNumber + 1;
          }
        }
      }

      const formattedNumber = nextNumber.toString().padStart(5, "0");

      return `CPX-PYMT/${fy}/${formattedNumber}`;
    } catch (error) {
      console.error("Capex ID Error:", error);
      return `CPX-PYMT/${getFinancialYear()}/00001`;
    }
  };
const uploadAttachments = async (capexId: string) => {
    try {
      if (!attachments || attachments.length === 0) return;

      const safeCapexId = capexId.replace(/\//g, "_");

      const libraryName = "CapexPaymentDocs";
      const webUrl = context.pageContext.web.serverRelativeUrl;

      const folderPath = `${webUrl}/${libraryName}/${safeCapexId}`;

      await sp.web.folders.addUsingPath(`${libraryName}/${safeCapexId}`);

      for (const file of attachments) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }

      console.log("Files uploaded successfully");
    } catch (error) {
      console.error("❌ Upload error:", error);
    }
  };
  // const uploadAttachments = async (capexId: string) => {
  //   try {
  //     if (!attachments || attachments.length === 0) return;

  //     const safeCapexId = capexId.replace(/\//g, "_");

  //     const libraryName = "CapexPaymentDocs";
  //     const webUrl = context.pageContext.web.serverRelativeUrl;

  //     const folderPath = `${webUrl}/${libraryName}/${safeCapexId}`;

  //     // ✅ Ensure folder
  //     await sp.web.folders.addUsingPath(`${libraryName}/${safeCapexId}`);

  //     // ✅ Upload files properly
  //     for (const file of attachments) {
  //       await sp.web
  //         .getFolderByServerRelativePath(folderPath)
  //         .files.addUsingPath(file.name, file, { Overwrite: true });
  //     }

  //     console.log("✅ Files uploaded successfully");
  //   } catch (error) {
  //     console.error("❌ Upload error:", error);
  //   }
  // };
  // 
  const buildApprovalPreview = async (employee: any) => {
    const flow: any[] = [];

    // RM
    if (employee.ReportingManager?.Title) {
      flow.push({
        Name: employee.ReportingManager.Title,
        Role: "RM",
        Status: "Pending",
      });
    }

    // HOD
    if (employee.HOD?.Title) {
      flow.push({
        Name: employee.HOD.Title,
        Role: "HOD",
        Status: "Pending",
      });
    }

    // 🔥 Matrix from list
    const matrixData = await sp.web.lists
      .getByTitle("CapexPaymentApprovalMatrix")
      .items.select("Role/RoleName,Approver/Title")
      .expand("Approver,Role")
      .filter("Status eq 'Active'")
      .orderBy("Level", true)();

    const matrixApprovers = matrixData.map((item: any) => ({
      Name: item.Approver?.Title,
      Role: item.Role?.RoleName,
      Status: "Pending",
    }));

    setApprovalMatrix([...flow, ...matrixApprovers]);
  };

  const buildApprovalFlow = async () => {
    const flow: any[] = [];

    // 🔹 RM
    if (employee.ReportingManager?.Id) {
      flow.push({
        Id: employee.ReportingManager.Id,
        Name: employee.ReportingManager.Title,
        Role: "RM",
        Level: 1,
        Status: "Pending",
      });
    }

    // 🔹 HOD
    if (employee.HOD?.Id) {
      flow.push({
        Id: employee.HOD.Id,
        Name: employee.HOD.Title,
        Role: "HOD",
        Level: 2,
        Status: "Pending",
      });
    }

    // 🔹 Matrix approvers
    const matrixData = await sp.web.lists
      .getByTitle("CapexPaymentApprovalMatrix")
      .items.select("Role/RoleName,Approver/Id,Approver/Title,Level/Level")
      .expand("Approver,Role,Level")
      .filter("Status eq 'Active'")
      .orderBy("Level", true)();

    const matrixApprovers = matrixData.map((item: any, index: number) => ({
      Id: item.Approver?.Id,
      Name: item.Approver?.Title,
      Role: item.Role?.RoleName,
      Level: flow.length + index + 1,
      Status: "Pending",
    }));

    const finalFlow = [...flow, ...matrixApprovers];

    // 🔥 first approver active
    if (finalFlow.length > 0) {
      finalFlow[0].Status = "In Progress";
    }

    return finalFlow;
  };

  const validateForm = () => {
    const errors: string[] = [];

    if (!selectedVendorId || selectedVendorId === 0) {
      errors.push("Please select Vendor");
    }

    if (!poNumber || poNumber.trim() === "") {
      errors.push("Please enter PO Number");
    }

    if (!poDate || poDate === "Invalid Date") {
      errors.push("Please select PO Date");
    }

    if (poDate > localDate) {
      errors.push("PO date cannot be a future date");
      // return;
    }

    if (!poTerms || poTerms.trim() === "") {
      errors.push("Please enter PO Terms");
    }

    if (!poAmount || poAmount.trim() === "") {
      errors.push("Please enter PO Amount");
    }

    // 🔥 MRN Validation
    if (!mrnNumber || mrnNumber.trim() === "") {
      errors.push("Please enter MRN Number");
    }

    if (!mrnDate || mrnDate === "Invalid Date") {
      errors.push("Please select MRN Date");
    }
     if (mrnDate > localDate) {
      errors.push("MRN date cannot be a future date");
      // return;
    }

    if (!mrnAmount || mrnAmount.trim() === "") {
      errors.push("Please enter MRN Amount");
    }

    if (!requestedAmount || requestedAmount.trim() === "") {
      errors.push("Please enter Requested Amount");
    }

    // 🔥 Final Payment Validation
   

    if (finalPayment === "Yes" && !installationDetails) {
      errors.push("Please enter Installation Details");
    }

    if (!attachments || attachments.length === 0) {
      errors.push("Please upload attachment");
    }
    if (!requestedAmount || requestedAmount.trim() === "") {
      errors.push("Please enter Requested Amount");
    }

    if (
      requestedAmount &&
      mrnAmount &&
      Number(requestedAmount) > Number(mrnAmount)
    ) {
      errors.push(
        "Requested Amount for Payment should not be greater than MRN Amount (GST)"
      );
    }
    return errors;
  };

  const handleSubmit = async () => {
    debugger;
   if (submitRef.current) return;

    //submitRef.current = true;
    setIsSubmitting(true);

    try {
      const errors = validateForm();

      if (errors.length > 0) {
        alert(errors.join("\n")); // 👈 shows exactly like your screenshot
        return;
      }

      const capexId = await generateCapexId();

      // ✅ Validate Vendor

      // ✅ Get Email from PeoplePicker
      // const userEmail = selectedUser[0]?.secondaryText;

      // if (!userEmail) {
      //   alert("User email not found");
      //   return;
      // }

      // ✅ Ensure User (FIX ERROR)
      // const ensuredUser = await sp.web.ensureUser(userEmail);
      const flow = await buildApprovalFlow();

      const currentApprover = flow.length > 0 ? flow[0].Id : null;

      // 🔥 workflow history
      const WorkflowHistory = [
        {
          CurrentApprover: employee.EmployeeName,
          ActionTaken: "Submitted",
          Comment: remarks,
          Date: new Date().toISOString(),
        },
      ];
      await sp.web.lists.getByTitle("CapexPayment").items.add({
        Title: capexId,
        CapexId: capexId,

        // Employee
        EmployeeCode: employee.EmployeeCode,
        EmployeeName: employee.EmployeeName,
        Division: employee.Division,
        Location: employee.Location,
        Email: employee.EmployeeEmail,
        RM: employee.ReportingManager?.Title,
        HOD: employee.HOD?.Title,
        ContactNo: employee.ContactNo,
        EmployeeStatus: employee.EmployeeStatus,

        // Vendor
        VendorCode: selectedVendorCode,

        VendorName: selectedVendorName,

        // PO
        PONumber: poNumber,
        PODate: poDate ? new Date(poDate) : null,
        POPaymentTerms: poTerms,
        POAmount: poAmount ? poAmount.toString() : "",

        // 🔥 NEW MRN FIELDS
        MRNNumber: mrnNumber,
        MRNDtae: mrnDate ? new Date(mrnDate) : null,

        MRNAmountwithGST: mrnAmount?.toString(),
        RequestedAmountforPayment: requestedAmount
          ? requestedAmount.toString()
          : "",

        // 🔥 FINAL PAYMENT
        FinalPaymentAgainstPO: finalPayment === "Yes",

        InstallationDetails: installationDetails,

        // Advance
        // RequestAdvanceAmount: advanceAmount,
        // PaidAmount: paidAmount,

        // PIC

        StatusFlow: "Pending for Approver",
        Status: "Pending for Approver",
        ApprovalMatrix: JSON.stringify(flow),
        CurrentApproverId: currentApprover,
        WorkflowHistory: JSON.stringify(WorkflowHistory),
      });

      debugger;
      await uploadAttachments(capexId); // 🔥 FIXED

      console.log("Attachments:", attachments);
      alert("Submitted successfully ✅");

      // 🔥 REDIRECT
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=User";
    } catch (error: any) {
      console.error("FULL ERROR:", error);
      console.error("SP ERROR:", error?.data?.responseBody);
      alert(error?.data?.responseBody || "Error while saving ❌");
    }
    finally {
      setIsSubmitting(false);
    }
  };

  const handledraft = async () => {
   
 if (draftRef.current) return;
    setIsDraftSaving(true);

    try {
      const capexId = await generateCapexId();

      let ensuredUserId: number | null = null;

      // ✅ Only process if user selected
      if (selectedUser && selectedUser.length > 0) {
        const userEmail = selectedUser[0]?.secondaryText;

        if (userEmail) {
          const ensuredUser = await sp.web.ensureUser(userEmail);
          ensuredUserId = ensuredUser.Id;
        }
      }
      const flow = await buildApprovalFlow();

      // 🔥 ALL SHOULD BE PENDING (NO IN PROGRESS)
      flow.forEach((f: any) => (f.Status = "Pending"));
      const currentApprover = flow.length > 0 ? flow[0].Id : null;

      // =========================
      // 🔥 WORKFLOW HISTORY (DRAFT)
      // =========================
      const WorkflowHistory = [
        {
          CurrentApprover: employee.EmployeeName,
          ActionTaken: "Draft Saved",
          Comment: remarks || "",
          Date: new Date().toISOString(),
        },
      ];
      await sp.web.lists.getByTitle("CapexPayment").items.add({
        Title: capexId,
        CapexId: capexId,

        // Employee
        EmployeeCode: employee.EmployeeCode,
        EmployeeName: employee.EmployeeName,
        Division: employee.Division,
        Location: employee.Location,
        Email: employee.EmployeeEmail,
        RM: employee.ReportingManager?.Title,
        HOD: employee.HOD?.Title,
        ContactNo: employee.ContactNo,
        EmployeeStatus: employee.EmployeeStatus,

        // Vendor
        VendorCode: selectedVendorCode,

        VendorName: selectedVendorName,

        // PO
        PONumber: poNumber,
        PODate: poDate ? new Date(poDate) : null,
        POPaymentTerms: poTerms,
        POAmount: poAmount ? poAmount.toString() : "",

        // 🔥 NEW MRN FIELDS
        MRNNumber: mrnNumber,
        MRNDtae: mrnDate ? new Date(mrnDate) : null,

        MRNAmountwithGST: mrnAmount?.toString(),
        RequestedAmountforPayment: requestedAmount
          ? requestedAmount.toString()
          : "",

        // 🔥 FINAL PAYMENT
        FinalPaymentAgainstPO: finalPayment === "Yes",

        InstallationDetails: installationDetails,

        // Advance
        // RequestAdvanceAmount: advanceAmount,
        // PaidAmount: paidAmount,

        // PIC

        StatusFlow: "Draft",
        Status: "Draft",
         ApprovalMatrix: JSON.stringify(flow),
         CurrentApproverId: currentApprover,
         WorkflowHistory: JSON.stringify(WorkflowHistory),
      });

      const safeCapexId = capexId.replace(/\//g, "_");
      void uploadAttachments(safeCapexId);


      

      alert("Draft saved successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=User";
    } catch (error) {
      console.error("ERROR:", error);
      alert("Error while saving ❌");
    }
    finally {
      draftRef.current = false;
      setIsDraftSaving(false);
    }
  };

  React.useEffect(() => {
    if (!context) return;

    void getLoggedInUser();
    void getVendors(); // 👈 ADD THIS
    //  buildApprovalPreview();
  }, [context]);

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1> Capex Payment Request </h1>
            </div>
           {approvalMatrix.length === 0 ? (
              <p>Loading...</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                   <li className={`approval-step`}>
                      {`Initiator`} - {employee.EmployeeName}
                    </li>
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${index === 0 ? "active" : ""}`}
                    >
                      {a.Role} - {a.Name}
                    </li>
                  ))}
                </ul>
              </div>
            )}
            <div className="borderedbox">
              <div className="heading1">
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Employee Code" className="font">
                      Employee Code
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">
                      Employee Name{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">
                      Employee Email{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.EmployeeEmail}
                    </label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">
                      Contact No
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">
                      Employee Status
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">
                      Division
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">
                      Location
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">
                      RM
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.ReportingManager?.Title}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">
                      HOD
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.HOD?.Title}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor & PO Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Code&nbsp; <span className="required">*</span> </label>
                    <select
                      value={selectedVendorId || ""}
                      onChange={(e) => {
                        const id = Number(e.target.value);

                        const vendor = vendors.find((v) => v.Id === id);

                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                        setSelectedVendorCode(vendor?.VendorCode || "");

                      
                         if (id > 0) {
                            void getPreviousAdvances(id);
                          } else {
                            // Clear table data
                            setPreviousAdvances([]);
                          }
                      }}
                      className="formtext-control"
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <input
                      value={selectedVendorName}
                      className="form-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number&nbsp; <span className="required">*</span> </label>
                    <input
                      value={poNumber}
                      className="form-control"
                      onChange={(e) => setPoNumber(e.target.value)}
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date&nbsp; <span className="required">*</span> </label>
                    <input
                      type="date"
                      value={poDate}
                      className="form-control"
                       max={new Date().toISOString().split("T")[0]} 
                      onChange={(e) => setPoDate(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (GST)&nbsp; <span className="required">*</span> </label>
                    <input
                      value={poAmount}
                      className="form-control"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoAmount)
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Payment Terms &nbsp; <span className="required">*</span> </label>
                    <input
                      value={poTerms}
                      className="form-control"
                      onChange={(e) => setPoTerms(e.target.value)}
                    />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Advance Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Number&nbsp; <span className="required">*</span> </label>
                    <input
                      value={mrnNumber}
                      onChange={(e) => setMrnNumber(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Date&nbsp; <span className="required">*</span> </label>
                    <input
                      type="date"
                      value={mrnDate}
                       max={new Date().toISOString().split("T")[0]}
                      onChange={(e) => setMrnDate(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Amount (GST)&nbsp; <span className="required">*</span> </label>
                    <input
                      value={mrnAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setMrnAmount)
                      }
                      className="form-control"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Requested Amount for Payment &nbsp; <span className="required">*</span> </label>
                    <input
                      value={requestedAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setRequestedAmount)
                      }
                      className="form-control"
                    />
                  </div>
                </div>
              </div>
              <div className="main-formcontainer" style={{ marginTop: "10px" }}>
                <div className="row mb-20">
                  <div className='col-md-4'>
                    <label className='font fontblock'>Whether this is the Final Payment against the PO &nbsp; <span className="required">*</span> </label>
                    <div className='radioalign'>
                      <label className='fonttext'>
                        <div>
                          <input
                            type="radio"
                            value="Yes"
                            checked={finalPayment === "Yes"}
                            onChange={(e) => setFinalPayment(e.target.value)}
                          />
                        </div>
                        &nbsp;
                        <div>Yes</div>
                      </label>
                      &nbsp;&nbsp;
                      <label className='fonttext '>
                        <div>
                          <input
                            type="radio"
                            value="No"
                            checked={finalPayment === "No"}
                            onChange={(e) => setFinalPayment(e.target.value)}
                          />
                        </div>
                        &nbsp;
                        <div>No</div>
                      </label>
                    </div>
                  </div>
                  {finalPayment === "Yes" && (
                    <div className="col-md-4">
                      <label className="font">Installation Details&nbsp; <span className="required">*</span> </label>
                      <textarea
                        value={installationDetails}
                        onChange={(e) =>
                          setInstallationDetails(e.target.value)
                        }
                        className="form-control"
                      />
                    </div>
                  )}
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Previous Advances</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <div className="table-vert-scroll">
                        <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                          <thead
                            className="text-white"
                            style={{ backgroundColor: "rgb(60, 62, 69)" }}
                          >
                            <tr>
                              <th className="px-4 py-2">PO Number</th>
                              <th className="px-4 py-2">Previous Advance</th>
                              <th className="px-4 py-2">Requested Date</th>
                              <th className="px-4 py-2">Paid Date</th>
                              <th className="px-4 py-2">MRN No</th>
                              <th className="px-4 py-2">Settled Amount</th>
                              <th className="px-4 py-2">Pending Advance</th>
                            </tr>
                          </thead>
                          <tbody>
                            {previousAdvances.length === 0 ? (
                              <tr>
                                <td colSpan={7} style={{ textAlign: "center" }}>
                                 No previous advances available
                                </td>
                              </tr>
                            ) : (
                              previousAdvances.map(
                                (item: any, index: number) => {
                                  const pending = Math.max(
                                    0,
                                    Number(item.RequestAdvanceAmount || 0) -
                                    Number(item.PaidAmount || 0),
                                  );
                                  return (
                                    <tr key={index}>
                                      <td>{item.PONumber}</td>
                                      <td>{item.RequestAdvanceAmount}</td>

                                      <td>
                                        {item.Created
                                          ? new Date(
                                            item.Created,
                                          ).toLocaleDateString()
                                          : ""}
                                      </td>

                                      <td>
                                        {item.VoucherDate
                                          ? new Date(
                                            item.VoucherDate,
                                          ).toLocaleDateString()
                                          : ""}
                                      </td>

                                      <td>{item.VoucherNumber}</td>
                                      <td>{item.PaidAmount}</td>
                                      <td>{pending}</td>
                                    </tr>
                                  );
                                },
                              )
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              <div className="main-formcontainer" style={{ marginTop: "10px" }}>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attach</label>{" "}
                    <span className="required" style={{ color: "red" }}>
                      *
                    </span>
                    <input
                      type="file"
                      multiple
                      className="form-control"
                      onChange={(e) => {
                        if (e.target.files) {
                          setAttachments((prev) => [
                            ...prev,
                            ...Array.from(e.target.files),
                          ]);
                        }
                      }}
                    />
                    {attachments.length > 0 && (
                      <ul style={{ marginTop: "10px" }}>
                        {attachments.map((file: File, index: number) => (
                          <li key={index}>
                            <a
                              href={URL.createObjectURL(file)}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {file.name}
                            </a>

                            <button
                              type="button"
                              style={{
                                marginLeft: "10px",
                                color: "red",
                                cursor: "pointer",
                              }}
                              onClick={() => handleRemoveFile(index)}
                            >
                              Remove
                            </button>
                          </li>
                        ))}
                      </ul>
                    )}
                    {/* {attachments.length > 0 && (
                      <ul style={{ marginTop: "10px" }}>
                        {attachments.map((file, index) => (
                          <li key={index}>
                            {file.name}

                            <button
                              type="button"
                              style={{
                                marginLeft: "10px",
                                color: "red",
                                cursor: "pointer",
                              }}
                              onClick={() => handleRemoveFile(index)}
                            >
                              Remove
                            </button>
                          </li>
                        ))}
                      </ul>
                    )} */}
                  </div>
                </div>
              </div>
              <div className='row my-3' >
                <div className='col-md-12'>
                  <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: "10px" }}>
                    <button
                      type="button"
                      onClick={!isSubmitting ? handleSubmit : undefined}
                      disabled={isSubmitting}
                      className="submit-btn"
                      style={{
                        pointerEvents: isSubmitting ? "none" : "auto",
                        opacity: isSubmitting ? 0.6 : 1,
                        cursor: isSubmitting ? "not-allowed" : "pointer",
                      }}
                    >
                      {isSubmitting ? "Submitting..." : "Submit"}
                    </button>
                    <button
                      type="button"
                      onClick={!isDraftSaving ? handledraft : undefined}
                      disabled={isDraftSaving}
                      className="Rework-btn"
                      style={{
                        pointerEvents: isDraftSaving ? "none" : "auto",
                        opacity: isDraftSaving ? 0.6 : 1,
                        cursor: isDraftSaving ? "not-allowed" : "pointer",
                      }}
                    >
                      {isDraftSaving ? "Saving..." : "Save as Draft"}
                    </button>
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

export default NewAdvanceform;
