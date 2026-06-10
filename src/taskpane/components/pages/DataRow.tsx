
import React from "react";
import { Box, Typography } from "@mui/material";

interface DataRowProps {
  icon: React.ReactNode;
  label: string;
  value: string;
  subValue?: string;
  isTitle?: boolean;
}

const DataRow: React.FC<DataRowProps> = ({ icon, label, value, subValue, isTitle }) => (
  <Box sx={{ display: 'flex', alignItems: 'flex-start', gap: 1.5 }}>
    <Box sx={{ mt: 0.5, color: 'primary.main', display: 'flex' }}>
      {icon}
    </Box>
    <Box sx={{ overflow: 'hidden', flexGrow: 1 }}>
      <Typography 
        variant="caption" 
        sx={{ 
          color: 'text.disabled', 
          fontWeight: 900, 
          fontSize: '0.55rem', 
          display: 'block', 
          textTransform: 'uppercase', 
          letterSpacing: 1 
        }}
      >
        {label}
      </Typography>
      
      <Typography 
        variant="body2" 
        sx={{ 
          fontWeight: isTitle ? 900 : 600, 
          color: '#1e293b', 
          fontSize: isTitle ? '0.9rem' : '0.8rem', 
          lineHeight: 1.3, 
          whiteSpace: 'nowrap', 
          overflow: 'hidden', 
          textOverflow: 'ellipsis' 
        }}
      >
        {value}
      </Typography>

      {subValue && (
        <Typography 
          variant="caption" 
          sx={{ 
            color: 'text.secondary', 
            display: 'block', 
            mt: -0.2, 
            whiteSpace: 'nowrap', 
            overflow: 'hidden', 
            textOverflow: 'ellipsis' 
          }}
        >
          {subValue}
        </Typography>
      )}
    </Box>
  </Box>
);

export default DataRow;