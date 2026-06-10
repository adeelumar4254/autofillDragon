// import React from "react";
// import { Box, IconButton, Tooltip } from "@mui/material";
// import { Logout as LogoutIcon } from "@mui/icons-material";

// interface ProfileHeaderProps {
//   onLogout?: () => void;
// }

// const ProfileHeader: React.FC<ProfileHeaderProps> = ({ onLogout }) => {
//   return (
//     <Box
//       sx={{
//         display: "flex",
//         justifyContent: "space-between",
//         alignItems: "center",
//         marginTop: "-10px", 
//         px: 1,
//       }}
//     >
    
//       <Box />

//       {/* --- RIGHT: LOGOUT ICON ONLY --- */}
//       <Tooltip title="Logout" arrow>
//         <IconButton
//           onClick={onLogout}
//           size="small"
//           sx={{
//             p: 0,
//             color: "#1f5975", 
//             transition: "all 0.2s ease",
//             "&:hover": {
//               transform: "scale(1.1)",
//               background: "transparent",
//               color: "#d32f2f", // Darker red on hover
//             },
//           }}
//         >
//           <LogoutIcon sx={{ fontSize: 25 }} /> 
    
//         </IconButton>
//       </Tooltip>
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
  ListItemIcon, Divider
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

  // Keep userName for the menu display, but we don't need userInitial anymore
  const userName = typeof Office !== "undefined" ? Office.context.mailbox?.userProfile?.displayName : "Outlook User";

  return (
    <>
      <Box sx={{
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        marginTop: "5px",
        px: 1 
      }}>
        
        {/* --- TITLE SECTION --- */}
        <Typography variant="h6" sx={{ fontWeight: 800, color: "#1f5975", fontSize: "1rem" }}>
          {title || "Email Summary"}
        </Typography>

        {/* --- PROFILE ICON BUTTON --- */}
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
            {/* Always show the AccountCircleIcon */}
            <AccountCircleIcon sx={{ fontSize: 26, color: "#fff" }} />
          </Avatar>
        </IconButton>
      </Box>

      {/* --- PROFILE MENU --- */}
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
              minWidth: 200,
              background: "rgba(255, 255, 255, 0.95)",
              backdropFilter: "blur(20px)",
              border: "1px solid rgba(255,255,255,0.5)",
              boxShadow: '0px 10px 25px rgba(0,0,0,0.1)',
              '& .MuiMenuItem-root': {
                fontSize: '0.85rem',
                fontWeight: 600,
                py: 1.2,
                borderRadius: '8px',
                mx: 0.5,
                transition: '0.2s',
              }
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

        <MenuItem onClick={handleInternalLogout} sx={{ color: 'error.main', '&:hover': { bgcolor: 'rgba(211, 47, 47, 0.05)' } }}>
          <ListItemIcon>
            <LogoutIcon fontSize="small" sx={{ color: 'error.main' }} />
          </ListItemIcon>
          Logout
        </MenuItem>
      </Menu>
    </>
  );
};

export default ProfileHeader;