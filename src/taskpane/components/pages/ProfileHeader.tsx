// import React from "react";
// import { Box, Button, Typography } from "@mui/material";
// import { Logout as LogoutIcon } from "@mui/icons-material";

// interface ProfileHeaderProps {
//   onLogout?: () => void;
//   title?: string; 
// }

// const ProfileHeader: React.FC<ProfileHeaderProps> = ({ onLogout, title }) => {
//   return (
//     <Box
//       sx={{
//         display: "flex",
//         justifyContent: "space-between",
//         alignItems: "center",
//         marginTop: "10px",
//         px: 1,
//         mb: 1
//       }}
//     >
//       <Typography 
//         variant="h6" 
//         sx={{ 
//           fontSize: "1rem", // Slightly smaller font for balance
//           fontWeight: 800, 
//           color: "#1f5975" 
//         }}
//       >
//         {title}
//       </Typography>

//       <Button
//         onClick={onLogout}
//         variant="outlined"
//         size="small"
//         startIcon={<LogoutIcon sx={{ fontSize: 16 }} />} // Smaller icon (16 instead of 18)
//         sx={{
//           color: "#ff4d4d",
//           borderColor: "#ff4d4d",
//           borderRadius: "50px",
//           textTransform: "none",
//           fontWeight: 700,
//           fontSize: "0.7rem",
//           minWidth: "auto",   
//           px: 1,            
//           py: 0.1,            
//           "& .MuiButton-startIcon": {
//             marginRight: "4px" 
//           },
//           "&:hover": {
//             borderColor: "#d32f2f",
//             backgroundColor: "rgba(211, 47, 47, 0.04)",
//             color: "#d32f2f",
//           },
//         }}
//       >
//         Logout
//       </Button>
//     </Box>
//   );
// };

// export default ProfileHeader;









// import React, { useState } from "react";
// import {
//   Box, Typography, IconButton, Avatar, Menu, MenuItem,
//   ListItemIcon, Divider
// } from "@mui/material";
// import {
//   Logout as LogoutIcon,
//   AccountCircle as AccountCircleIcon
// } from "@mui/icons-material";


// interface ProfileHeaderProps {
//   title?: string; // Made optional since you are using a logo instead
//   onLogout?: () => void;
// }

// const ProfileHeader: React.FC<ProfileHeaderProps> = ({ title, onLogout }) => {
//   const [anchorEl, setAnchorEl] = useState<null | HTMLElement>(null);
//   const openMenu = Boolean(anchorEl);

//   const handleProfileClick = (event: React.MouseEvent<HTMLElement>) => {
//     setAnchorEl(event.currentTarget);
//   };

//   const handleCloseMenu = () => {
//     setAnchorEl(null);
//   };

//   const handleInternalLogout = () => {
//     handleCloseMenu();
//     if (onLogout) onLogout();
//   };

//   const userName = typeof Office !== "undefined" ? Office.context.mailbox?.userProfile?.displayName : "Outlook User";
//   const userInitial = userName?.charAt(0) || "";
// console.log(title)
//   return (
//     <>
//       <Box sx={{
//         display: 'flex',
//         justifyContent: 'space-between',
//         alignItems: 'center',
//         marginTop: "5px", // Adjusted margin for better spacing
//         px: 1 // Added small horizontal padding
//       }}>
        
//         {/* --- LOGO SECTION --- */}
//         <Typography variant="h6" sx={{ fontWeight: 800, color: "#1f5975", fontSize: "1rem" }}>
//           {title || "Email Summary"}
//         </Typography>

//         {/* --- PROFILE SECTION --- */}
//         <IconButton
//           onClick={handleProfileClick}
//           size="small"
//           disableRipple
//           sx={{
//             p: 0,
//             background: "transparent",
//             border: "none",
//             boxShadow: "none",
//             transition: "all 0.2s ease",
//             "&:hover": { 
//               transform: "scale(1.1)", 
//               background: "transparent"
//             }
//           }}
//         >
//           <Avatar 
//             sx={{ 
//               width: 32, 
//               height: 32, 
//               bgcolor: "#6aa2ba", 
//               fontSize: '0.85rem', 
//               fontWeight: 700,
//               boxShadow: "0 2px 8px rgba(106, 162, 186, 0.3)" 
//             }}
//           >
//             {userInitial || <AccountCircleIcon sx={{ fontSize: 20 }} />}
//           </Avatar>
//         </IconButton>
//       </Box>

//       {/* --- PROFILE MENU --- */}
//       <Menu
//         anchorEl={anchorEl}
//         open={openMenu}
//         onClose={handleCloseMenu}
//         transformOrigin={{ horizontal: 'right', vertical: 'top' }}
//         anchorOrigin={{ horizontal: 'right', vertical: 'bottom' }}
//         slotProps={{
//           paper: {
//             elevation: 0,
//             sx: {
//               mt: 1.5,
//               borderRadius: '12px',
//               minWidth: 200,
//               background: "rgba(255, 255, 255, 0.95)",
//               backdropFilter: "blur(20px)",
//               border: "1px solid rgba(255,255,255,0.5)",
//               boxShadow: '0px 10px 25px rgba(0,0,0,0.1)',
//               '& .MuiMenuItem-root': {
//                 fontSize: '0.85rem',
//                 fontWeight: 600,
//                 py: 1.2,
//                 borderRadius: '8px',
//                 mx: 0.5,
//                 transition: '0.2s',
//               }
//             },
//           }
//         }}
//       >
//         <Box sx={{ px: 2, py: 1.5 }}>
//           <Typography variant="caption" sx={{ color: 'text.secondary', display: 'block', fontWeight: 700, textTransform: 'uppercase', mb: 0.5, letterSpacing: 1 }}>
//             Logged in as
//           </Typography>
//           <Typography variant="body2" sx={{ fontWeight: 800, color: '#1e293b' }}>
//             {userName}
//           </Typography>
//         </Box>

//         <Divider sx={{ my: 0.5, opacity: 0.5 }} />

//         <MenuItem onClick={handleInternalLogout} sx={{ color: 'error.main', '&:hover': { bgcolor: 'rgba(211, 47, 47, 0.05)' } }}>
//           <ListItemIcon>
//             <LogoutIcon fontSize="small" sx={{ color: 'error.main' }} />
//           </ListItemIcon>
//           Logout
//         </MenuItem>
//       </Menu>
//     </>
//   );
// };

// export default ProfileHeader;







import React, { useState } from "react";
import {
  Box, Typography, IconButton, Avatar, Menu, MenuItem,
  ListItemIcon, Divider, Button
} from "@mui/material";
import {
  Logout as LogoutIcon,
  AccountCircle as AccountCircleIcon
} from "@mui/icons-material";

interface ProfileHeaderProps {
  title?: string;
  onLogout?: () => void;
}

const ProfileHeader: React.FC<ProfileHeaderProps> = ({ title, onLogout }) => {
  const [anchorEl, setAnchorEl] = useState<null | HTMLElement>(null);
  const openMenu = Boolean(anchorEl);

  const handleProfileClick = (event: React.MouseEvent<HTMLElement>) => {
    setAnchorEl(event.currentTarget);
  };

  const handleCloseMenu = () => {
    setAnchorEl(null);
  };

  const handleInternalLogout = () => {
    handleCloseMenu();
    if (onLogout) onLogout();
  };

  const userName = typeof Office !== "undefined" ? Office.context.mailbox?.userProfile?.displayName : "Outlook User";
console.log(userName);
  return (
    <>
      <Box sx={{
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        marginTop: "5px",
        px: 1 
      }}>
        
        <Typography variant="h6" sx={{ fontWeight: 800, color: "#1f5975", fontSize: "1rem" }}>
          {title || "Email Summary"}
        </Typography>

        <IconButton
          onClick={handleProfileClick}
          size="small"
          disableRipple
          sx={{
            p: 0,
            background: "transparent",
            transition: "all 0.2s ease",
            "&:hover": { 
              transform: "scale(1.1)", 
              background: "transparent"
            }
          }}
        >
          <Avatar 
            sx={{ 
              width: 26, 
              height: 26, 
              bgcolor: "#6aa2ba", 
              boxShadow: "0 2px 8px rgba(106, 162, 186, 0.3)" 
            }}
          >
            <AccountCircleIcon sx={{ fontSize: 26, color: "#fff" }} />
          </Avatar>
        </IconButton>
      </Box>

      <Menu
        anchorEl={anchorEl}
        open={openMenu}
        onClose={handleCloseMenu}
        transformOrigin={{ horizontal: 'right', vertical: 'top' }}
        anchorOrigin={{ horizontal: 'right', vertical: 'bottom' }}
        slotProps={{
          paper: {
            elevation: 0,
            sx: {
              mt: 1.5,
              borderRadius: '12px',
              minWidth: 220, // Slightly wider to fit the pill button
              background: "rgba(255, 255, 255, 0.95)",
              backdropFilter: "blur(20px)",
              border: "1px solid rgba(255,255,255,0.5)",
              boxShadow: '0px 10px 25px rgba(0,0,0,0.1)',
            },
          }
        }}
      >
        <Box sx={{ px: 2, py: 1.5 }}>
          <Typography variant="caption" sx={{ color: 'text.secondary', display: 'block', fontWeight: 700, textTransform: 'uppercase', mb: 0.5, letterSpacing: 1 }}>
            Logged in as
          </Typography>
          <Typography variant="body2" sx={{ fontWeight: 800, color: '#1e293b' }}>
            {userName}
          </Typography>
        </Box>

        <Divider sx={{ my: 0.5, opacity: 0.5 }} />

        {/* --- CUSTOM PILL LOGOUT BUTTON --- */}
       {/* --- SMALLER PILL LOGOUT BUTTON --- */}
        <Box sx={{ px: 2, py: 1.5, display: 'flex', justifyContent: 'center',float: "left" }}>
          <MenuItem 
            onClick={handleInternalLogout} 
            disableRipple 
            sx={{ 
              color: '#ff4d4d', 
              border: '1px solid #ff4d4d',
              borderRadius: '50px',
              width: 'fit-content',
              minWidth: '100px', 
              minHeight: 'auto',  
              px: 2,                
              py: 0.6,              
              fontSize: '0.75rem',  
              fontWeight: 700,
              transition: 'all 0.2s ease',
              '&:hover': { 
                bgcolor: 'rgba(255, 77, 77, 0.05)',
                borderColor: '#d32f2f',
                color: '#d32f2f',
                transform: 'translateY(-1px)'
              } 
            }}
          >
            <LogoutIcon sx={{ mr: 1, fontSize: 16 }} /> {/* Smaller icon */}
            Logout
          </MenuItem>
        </Box>
      </Menu>
    </>
  );
};

export default ProfileHeader;