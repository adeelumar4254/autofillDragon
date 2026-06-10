import React from "react";
import { Box, IconButton, Tooltip } from "@mui/material";
import { Logout as LogoutIcon } from "@mui/icons-material";

interface ProfileHeaderProps {
  onLogout?: () => void;
}

const ProfileHeader: React.FC<ProfileHeaderProps> = ({ onLogout }) => {
  return (
    <Box
      sx={{
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        marginTop: "-10px", 
        px: 1,
      }}
    >
    
      <Box />

      {/* --- RIGHT: LOGOUT ICON ONLY --- */}
      <Tooltip title="Logout" arrow>
        <IconButton
          onClick={onLogout}
          size="small"
          sx={{
            p: 0,
            color: "#1f5975", 
            transition: "all 0.2s ease",
            "&:hover": {
              transform: "scale(1.1)",
              background: "transparent",
              color: "#d32f2f", // Darker red on hover
            },
          }}
        >
          <LogoutIcon sx={{ fontSize: 25 }} /> 
    
        </IconButton>
      </Tooltip>
    </Box>
  );
};

export default ProfileHeader;