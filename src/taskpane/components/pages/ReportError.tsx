// import React, { useState } from "react";
// import { Box, Typography, TextField, Button, IconButton, Paper } from "@mui/material";
// import { ArrowBack as BackIcon } from "@mui/icons-material";
// import { useNavigate, useLocation } from "react-router-dom";
// import { useSnackbar } from "../../context/SnackbarContext";

// const ReportError = () => {
//   const navigate = useNavigate();
//   const location = useLocation();
//   const { showSnackbar } = useSnackbar();


//   // const incomingError = location.state?.error || "No specific error data available.";
//   const incomingError = location.state?.error || "No specific error data available.";
//   const uniqueId = location.state?.uniqueId || "";

//   const [description, setDescription] = useState("");
//   const [submitting, setSubmitting] = useState(false);



//   const handleSubmit = async () => {
//   if (!description.trim()) {
//     showSnackbar("Please describe the issue", "warning");
//     return;
//   }

//   setSubmitting(true);

//   try {
//     const token = window.localStorage.getItem("token") || "";

//     // Requirement: Post as form-urlencoded
//     const formData = new URLSearchParams();
//     formData.append("UniqueId", uniqueId);
//     formData.append("Description", description);

//     const response = await fetch(
//       "https://dragon.bookingbuilder.com/AutofillService/api/Error/report-error",
//       {
//         method: "POST",
//         headers: {
//           "Authorization": `Bearer ${token}`,
//         },
//         body: formData,
//       }
//     );

//     if (response.ok) {
//       showSnackbar("Error report submitted successfully", "success");
//       navigate("/home");
//     } else if (response.status === 401) {
//       showSnackbar("Session expired. Please log in again.", "error");
//     } else {
//       const errorText = await response.text();
//       console.error("Server Error:", errorText);
//       throw new Error(`Server returned ${response.status}`);
//     }
//   } catch (error) {
//     console.error("Submission failed:", error);
//     showSnackbar("Failed to send report. Please try again.", "error");
//   } finally {
//     setSubmitting(false);
//   }
// };
//   // const handleSubmit = async () => {
//   //   if (!description.trim()) {
//   //     showSnackbar("Please describe the issue", "warning");
//   //     return;
//   //   }

//   //   setSubmitting(true);

//   //   try {
//   //     // Requirement: Post as form-urlencoded, not JSON
//   //     const token = window.localStorage.getItem("token") || "";

//   //     const formData = new URLSearchParams();
//   //     formData.append("UniqueId", uniqueId);
//   //     formData.append("Description", description);

//   //     const response = await fetch(
//   //       "https://dragon.bookingbuilder.com/AutofillService/api/Error/report-error",
//   //       {
//   //         method: "POST",
//   //         headers: {
//   //           "Content-Type": "application/x-www-form-urlencoded",
//   //           "Authorization": `Bearer ${token}`,
//   //         },
//   //         body: formData.toString(),
//   //       }
//   //     );

//   //     if (response.ok) {
//   //       showSnackbar("Error report submitted successfully", "success");
//   //       navigate("/home");
//   //     } else {
//   //       throw new Error("Failed to submit report");
//   //     }
//   //   } catch (error) {
//   //     showSnackbar("Failed to send report. Please try again.", "error");
//   //   } finally {
//   //     setSubmitting(false);
//   //   }
//   // };

//   return (
//     <Box sx={{ height: "100vh", bgcolor: "#ffffff", display: "flex", flexDirection: "column" }}>
//       {/* Header */}
//       <Box sx={{ p: 2, display: "flex", alignItems: "center", borderBottom: "1px solid #f1f5f9" }}>
//         <IconButton onClick={() => navigate("/home")} size="small" sx={{ mr: 1 }}>
//           <BackIcon fontSize="small" />
//         </IconButton>
//         <Typography variant="h6" sx={{ fontWeight: 700, fontSize: "1rem" }}>Report Issue</Typography>
//       </Box>

//       <Box sx={{ p: 2, flexGrow: 1, overflowY: 'auto' }}>
//         <Typography variant="caption" sx={{ fontWeight: 700, color: "error.main", mb: 0.5, display: 'block' }}>
//           DETECTED ERROR:
//         </Typography>
//         <Paper variant="outlined" sx={{ p: 1.5, mb: 1, bgcolor: '#fff5f5', borderColor: '#feb2b2', borderRadius: '8px' }}>
//           <Typography variant="body2" sx={{ fontFamily: 'monospace', color: '#c53030', wordBreak: 'break-all' }}>
//             {incomingError}
//           </Typography>
//         </Paper>

//         {/* Display Unique ID so user knows it's being tracked */}
//         <Typography variant="caption" sx={{ color: "text.secondary", mb: 3, display: 'block' }}>
//           ID: {uniqueId || "N/A"}
//         </Typography>

//         <Typography variant="body2" sx={{ color: "#1f5975", mb: 1, fontWeight: 500 }}>
//           Please describe what went wrong:
//         </Typography>
//         <TextField
//           fullWidth multiline rows={6}
//           placeholder="e.g. This is a confirmation for a hotel but it failed to parse the dates..."
//           value={description}
//           onChange={(e) => setDescription(e.target.value)}
//           sx={{
//             "& .MuiOutlinedInput-root": {
//               bgcolor: "#f8fafc", borderRadius: "12px",
//             }
//           }}
//         />
//       </Box>

//       <Box sx={{ p: 2, borderTop: "1px solid #f1f5f9" }}>
//         <Button
//           fullWidth size="small" variant="contained" disableElevation
//           onClick={handleSubmit} disabled={submitting}
//           sx={{ py: 1, bgcolor: "#1f5975", borderRadius: "10px", textTransform: "none", fontWeight: 700 }}
//         >
//           {submitting ? "Submitting..." : "Submit Report"}
//         </Button>
//       </Box>
//     </Box>
//   );
// };
// export default ReportError;


// src/components/ReportError/ReportError.tsx

import React, { useState } from "react";
import { Box, Typography, TextField, Button, IconButton, Paper } from "@mui/material";
import { ArrowBack as BackIcon } from "@mui/icons-material";
import { useNavigate, useLocation } from "react-router-dom";
import { useSnackbar } from "../../context/SnackbarContext";
import { reportErrorApi } from "../../services/ErrorReport"; // Import the new service

const ReportError = () => {
  const navigate = useNavigate();
  const location = useLocation();
  const { showSnackbar } = useSnackbar();

  // Retrieve data passed from the Home screen state
  const incomingError = location.state?.error || "No specific error data available.";
  const uniqueId = location.state?.uniqueId || "";

  const [description, setDescription] = useState("");
  const [submitting, setSubmitting] = useState(false);

  const handleSubmit = async () => {
    // 1. Validation
    if (!description.trim()) {
      showSnackbar("Please describe the issue before submitting", "warning");
      return;
    }

    setSubmitting(true);

    try {
    
      await reportErrorApi(uniqueId, description);

      // 3.Success Handling
      showSnackbar("Error report submitted successfully", "success");
      
      setTimeout(() => navigate("/home"), 500);

    } catch (error: any) {
      // 4. Error Handling
      console.log("Submission Error Details:", error);

      if (error.message === "UNAUTHORIZED") {
        showSnackbar("Your session has expired. Please log in again.", "error");
        navigate("/register"); 
      } else if (error.message.includes("No authentication token")) {
        showSnackbar("Authentication missing. Please log in.", "error");
        navigate("/register");
      } else {
        // Handle network errors or server-specific errors
        const errorMsg = error.message || "Failed to connect to the server.";
        showSnackbar(`Error: ${errorMsg}`, "error");
      }
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <Box sx={{ height: "100vh", bgcolor: "#ffffff", display: "flex", flexDirection: "column" }}>
      {/* Header */}
      <Box sx={{ p: 2, display: "flex", alignItems: "center", borderBottom: "1px solid #f1f5f9" }}>
        <IconButton onClick={() => navigate("/home")} size="small" sx={{ mr: 1 }}>
          <BackIcon fontSize="small" />
        </IconButton>
        <Typography variant="h6" sx={{ fontWeight: 700, fontSize: "1rem" }}>Cancel</Typography>
      </Box>

      {/* Content */}
      <Box sx={{ p: 2, flexGrow: 1, overflowY: 'auto' }}>
        <Typography variant="caption" sx={{ fontWeight: 700, color: "error.main", mb: 0.5, display: 'block' }}>
          TECHNICAL ERROR DETECTED:
        </Typography>
        
        <Paper variant="outlined" sx={{ p: 1.5, mb: 1, bgcolor: '#fff5f5', borderColor: '#feb2b2', borderRadius: '8px' }}>
          <Typography variant="body2" sx={{ fontFamily: 'monospace', color: '#c53030', wordBreak: 'break-all', fontSize: '0.75rem' }}>
            {incomingError}
          </Typography>
        </Paper>

        {/* <Typography variant="caption" sx={{ color: "text.secondary", mb: 3, display: 'block' }}>
          Tracking ID: {uniqueId || "Not Available"}
        </Typography> */}

        <Typography variant="body2" sx={{ color: "#1f5975", mb: 1, fontWeight: 500 }}>
          Additional Details:
        </Typography>
        
        <TextField
          fullWidth multiline rows={6}
          placeholder="Enter a detailed description of the issue..."
          value={description}
          onChange={(e) => setDescription(e.target.value)}
          disabled={submitting}
          sx={{
            "& .MuiOutlinedInput-root": {
              bgcolor: "#f8fafc", borderRadius: "12px",
            }
          }}
        />
      </Box>

      {/* Footer */}
      <Box sx={{ p: 2, borderTop: "1px solid #f1f5f9" }}>
        <Button
          fullWidth size="small" variant="contained" disableElevation
          onClick={handleSubmit} 
          disabled={submitting}
          sx={{ 
            py: 1, 
            bgcolor: "#1f5975", 
            borderRadius: "10px", 
            textTransform: "none", 
            fontWeight: 700,
            "&:hover": { bgcolor: "#154157" } 
          }}
        >
          {submitting ? "Submitting..." : "Submit Report"}
        </Button>
      </Box>
    </Box>
  );
};

export default ReportError;