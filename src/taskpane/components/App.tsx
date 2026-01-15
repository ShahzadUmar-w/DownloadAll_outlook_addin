// import React, { useEffect, useState, useCallback } from "react";
// import {
//   Container,
//   Typography,
//   Card,
//   CardContent,
//   List,
//   ListItem,
//   ListItemText,
//   Button,
//   Divider,
//   Box,
//   CircularProgress,
//   IconButton,
//   Tooltip,
// } from "@mui/material";
// import { ArrowDownload16Filled, ArrowDownload16Regular, Link12Filled } from "@fluentui/react-icons";
// // import DownloadIcon from '@mui/icons-material/Download';
// // import LinkOutlinedIcon from '@mui/icons-material/LinkOutlined';

// // --- Utility Functions (Strict Logic) ---

// const downloadableExtensions = [
//   ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".rar", ".csv", ".txt",
// ];

// // Target domain for strict filtering
// const TARGET_DOMAIN = "cdn.azams0.fds.api.mi-img.com";

// const isDownloadable = (url: string) => {
//     const lowerUrl = url.toLowerCase();

//     // 1. Strict Check: Agar URL target domain se nahi hai, toh reject kar dein
//     if (!lowerUrl.includes(TARGET_DOMAIN)) {
//         // Agar aap chahte hain ki TARGET_DOMAIN ke alawa koi aur standard file (.pdf) bhi na aaye, 
//         // toh sirf TARGET_DOMAIN check kaafi hai. Agar koi bhi standard file bhi chahiye, 
//         // toh niche wala logic use karein.
        
//         // **Current Goal:** Sirf specific domain ke links ya standard files.
        
//         // Standard file extension check (Query param safe)
//         const cleanUrl = url.split(/[?#]/)[0].toLowerCase();
//         return downloadableExtensions.some((ext) => cleanUrl.endsWith(ext));
//     }
    
//     // Target domain ko hamesha allow karo
//     return true;
// };

// const extractLinks = (html: string): string[] => {
//   const parser = new DOMParser();
//   const doc = parser.parseFromString(html, "text/html");
//   const anchors = Array.from(doc.querySelectorAll("a"));
  
//   let links = anchors.map((a) => a.href);

//   // Filtering: sirf http/https links jo downloadable hain
//   return links
//     .filter((href) => href.startsWith("http") && isDownloadable(href));
// };

// const processItemBody = (item: Office.MessageRead): Promise<string[]> => {
//     return new Promise((resolve, reject) => {
//       item.body.getAsync(Office.CoercionType.Html, (result) => {
//         if (result.status === Office.AsyncResultStatus.Succeeded) {
//           resolve(extractLinks(result.value));
//         } else {
//           console.error("Failed to get body:", result.error);
//           reject(result.error);
//         }
//       });
//     });
// };

// // --- React Component ---

// const App: React.FC = () => {
//   const [allLinks, setAllLinks] = useState<string[]>([]);
//   const [loading, setLoading] = useState<boolean>(true);

//   // Sequential Item Handler
//   const loadAndUnloadSingleItem = (itemId: string): Promise<string[]> => {
//     return new Promise((resolve) => {
      
//       Office.context.mailbox.loadItemByIdAsync(itemId, (loadResult) => {
//         if (loadResult.status !== Office.AsyncResultStatus.Succeeded) {
//           console.error(
//             `Failed to load item ${itemId} (Code: ${loadResult.error.code}): ${loadResult.error.message}`
//           );
//           return resolve([]);
//         }

//         const loadedItem:any = loadResult.value as unknown as Office.MessageRead;

//         processItemBody(loadedItem)
//           .then(resolve) 
//           .catch((e) => {
//             console.error(`Body processing failed for ${itemId}`, e);
//             resolve([]); 
//           })
//           .finally(() => {
//             loadedItem.unloadAsync(() => {
//               // Unload complete
//             });
//           });
//       });
//     });
//   };

//   // Main processing loop (Duplicates allowed)
//   const loadAndProcessSelectedItems = useCallback(() => {
//     setLoading(true);
//     setAllLinks([]);

//     Office.context.mailbox.getSelectedItemsAsync(async (asyncResult) => {
//       if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
//         setLoading(false);
//         return;
//       }

//       const selectedItems = asyncResult.value as Office.SelectedItemDetails[];
//       if (selectedItems.length === 0) {
//         setLoading(false);
//         return;
//       }
      
//       const itemIds = selectedItems.map(item => item.itemId);
//       let extractedLinks: string[] = [];

//       for (const itemId of itemIds) {
//         const linksFromItem = await loadAndUnloadSingleItem(itemId); 
        
//         // *** DUPLICATE LINKS ALLOWED: Directly push all links ***
//         linksFromItem.forEach(link => {
//             extractedLinks.push(link);
//         });
//       }

//       setAllLinks(extractedLinks);
//       setLoading(false);
//     });
//   }, []); 

//   // Initialization and Event Listener Setup (Same as before)
//   useEffect(() => {
//     Office.onReady((info) => {
//       if (info.host === Office.HostType.Outlook) {
//         if (Office.context.requirements.isSetSupported("Mailbox", "1.15")) {
//           Office.context.mailbox.addHandlerAsync(
//             Office.EventType.SelectedItemsChanged,
//             loadAndProcessSelectedItems,
//             (result) => {
//               if (result.status === Office.AsyncResultStatus.Succeeded) {
//                 loadAndProcessSelectedItems();
//               }
//             }
//           );
//         } else {
//           setLoading(false);
//         }
//       }
//     });
//   }, [loadAndProcessSelectedItems]);

//   const downloadAllLinks = () => {
//     allLinks.forEach((link) => {
//       const a = document.createElement("a");
//       a.href = link;
//       a.target = "_blank";
//       a.click();
//     });
//   };

//   // --- Attractive UI Rendering ---
//   return (
//     <Container maxWidth="sm" sx={{ p: 0, pt: 1 }}>
//       <Typography
//         variant="h5"
//         align="center"
//         sx={{ 
//             fontWeight: 'bold', 
//             color: '#106ebe', 
//             mb: 2, 
//             p: 1,
//             borderRadius: 1,
//             backgroundColor: '#f5f8ff'
//         }}
//       >
//         Multi-Email Downloader
//       </Typography>

//       <Card
//         sx={{ 
//           maxHeight: 350, 
//           overflowY: "auto", 
//           mb: 3, 
//           boxShadow: 3, 
//           borderRadius: 2,
//           border: '1px solid #e0e0e0'
//         }}
//       >
//         <CardContent sx={{ p: 0 }}>
//           {loading ? (
//             <Box display="flex" flexDirection="column" alignItems="center" p={4}>
//               <CircularProgress size={30} sx={{ mb: 1, color: '#106ebe' }} />
//               <Typography variant="body2" color="text.secondary">
//                 Analyzing {allLinks.length > 0 ? 'new' : 'selected'} messages...
//               </Typography>
//             </Box>
//           ) : allLinks.length === 0 ? (
//             <Typography variant="body1" color="text.secondary" align="center" p={3}>
//               No target downloadable files found.
//             </Typography>
//           ) : (
//             <List disablePadding>
//               {allLinks.map((link, idx) => (
//                 <React.Fragment key={idx}>
//                   <ListItem
//                     // key must be unique per item, using index is fine since we allow duplicates
//                     key={idx} 
//                     secondaryAction={
//                         <Tooltip title="Download this file">
//                             <IconButton 
//                                 edge="end" 
//                                 aria-label="download" 
//                                 href={link}
//                                 target="_blank"
//                                 size="small"
//                                 sx={{ color: '#106ebe' }}
//                             >
//                                 {/* <ArrowDownload16Filled fontSize="small"/> */}
//                             </IconButton>
//                         </Tooltip>
//                     }
//                     sx={{ 
//                       px: 2, 
//                       py: 0.5,
//                       backgroundColor: idx % 2 === 0 ? '#fcfcfc' : 'white', 
//                       "&:hover": { backgroundColor: "#e3f2fd" },
//                     }}
//                   >
//                     <ListItemText
//                       primary={
//                         <Box display="flex" alignItems="center">
//                             <Link12Filled style={{ fontSize: '0.9rem', marginLeft: 1, color: '#0d47a1' }} />
//                             {link}
//                         </Box>
//                       }
//                       primaryTypographyProps={{
//                         fontSize: "0.75rem",
//                         color: "#555",
//                         noWrap: true,
//                         sx: {
//                           overflow: "hidden",
//                           textOverflow: "ellipsis",
//                           fontFamily: "monospace",
//                         },
//                       }}
//                     />
//                   </ListItem>
//                   <Divider component="li" light />
//                 </React.Fragment>
//               ))}
//             </List>
//           )}
//         </CardContent>
//       </Card>

//       {allLinks.length > 0 && (
//         <Box textAlign="center">
//           <Button
//             variant="contained"
//             size="large"
//             onClick={downloadAllLinks}
//             disabled={loading}
//             startIcon={<ArrowDownload16Regular />}
//             sx={{
//               fontWeight: "bold",
//               px: 5,
//               py: 1.2,
//               fontSize: '1rem',
//               backgroundColor: "#2e7d32", // Green button for download
//               "&:hover": { backgroundColor: "#1b5e20" },
//               boxShadow: 4,
//             }}
//           >
//             Download All Files ({allLinks.length})
//           </Button>
//         </Box>
//       )}
//     </Container>
//   );
// };

// export default App;

import React, { useEffect, useState, useCallback } from "react";
import {
  Container,
  Typography,
  Card,
  CardContent,
  List,
  ListItem,
  ListItemText,
  Button,
  Divider,
  Box,
  CircularProgress,
  IconButton,
  Tooltip,
} from "@mui/material";
import { ArrowDownload16Regular, Link12Filled } from "@fluentui/react-icons";

// --- Configuration ---
const BATCH_SIZE = 10;
const DELAY_MS = 10000; // 10 seconds

// Target domain for strict filtering (Your Invoice Domain)
const TARGET_DOMAIN = "cdn.azams0.fds.api.mi-img.com";

// Common file extensions
const downloadableExtensions = [
  ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".rar", ".csv", ".txt", ".jpg", ".png",
];

// --- Utility Functions (Improved Filtering Logic) ---

/**
 * Checks if a URL should be included based on strict domain or file extension.
 * This ensures your specific invoice links are prioritized.
 */
const isDownloadable = (url: string): boolean => {
    const lowerUrl = url.toLowerCase();

    // 1. Strict Check: Agar URL target domain se hai, toh use TURANT allow kar dein.
    // Aapka invoice link isi check se pass hoga, irrespective of query params.
    if (lowerUrl.includes(TARGET_DOMAIN)) {
        return true;
    }
    
    // 2. Fallback Check: Agar target domain se nahi hai, toh dekho kya yeh koi common file extension hai?
    
    // Query parameters hatao
    const cleanUrl = lowerUrl.split(/[?#]/)[0];
    
    // Check extension
    return downloadableExtensions.some((ext) => cleanUrl.endsWith(ext));
};

const decodeHtml = (html: string): string => {
  const txt = document.createElement("textarea");
  txt.innerHTML = html;
  return txt.value;
};

const URL_REGEX =
  /(https?:\/\/[^\s"'<>]+(?:\.[^\s"'<>]+)+(?:\/[^\s"'<>]*)?)/gi;

const extractLinks = (html: string): string[] => {
  const decodedHtml = decodeHtml(html);

  const linksSet = new Set<string>();

  // 1. From anchor tags
  const parser = new DOMParser();
  const doc = parser.parseFromString(decodedHtml, "text/html");
  const anchors = Array.from(doc.querySelectorAll("a[href]"));

  anchors.forEach(a => {
    const href = a.getAttribute("href")?.trim();
    if (href && href.startsWith("http") && isDownloadable(href)) {
      linksSet.add(href);
    }
  });

  // 2. From plain text (VERY IMPORTANT)
  const textContent = doc.body.innerText || decodedHtml;
  const matches = textContent.match(URL_REGEX);

  if (matches) {
    matches.forEach(url => {
      const cleanUrl = url.trim();
      if (isDownloadable(cleanUrl)) {
        linksSet.add(cleanUrl);
      }
    });
  }

  return Array.from(linksSet);
};

// processItemBody and other utility functions remain the same
const processItemBody = (item: Office.MessageRead): Promise<string[]> => {
    return new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          // Note: getAsync fetches the entire body of the selected email/thread item
          resolve(extractLinks(result.value));
        } else {
          console.error("Failed to get body:", result.error);
          reject(result.error);
        }
      });
    });
};

// --- React Component ---

const App: React.FC = () => {
  const [allLinks, setAllLinks] = useState<string[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [currentBatch, setCurrentBatch] = useState<number>(0);


  // Ensures entire body is loaded, processed, and unloaded
  const loadAndUnloadSingleItem = (itemId: string): Promise<string[]> => {
    return new Promise((resolve) => {
      
      Office.context.mailbox.loadItemByIdAsync(itemId, (loadResult) => {
        if (loadResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(
            `Failed to load item ${itemId} (Code: ${loadResult.error.code})`
          );
          return resolve([]);
        }

        const loadedItem: any = loadResult.value as unknown as Office.MessageRead;

        processItemBody(loadedItem)
          .then(resolve) 
          .catch((e) => {
            console.error(`Body processing failed for ${itemId}`, e);
            resolve([]); 
          })
          .finally(() => {
            // Memory management for Outlook add-ins
            if (loadedItem.unloadAsync) {
                loadedItem.unloadAsync(() => {}); 
            }
          });
      });
    });
  };

  // Handles multiple selected items (threads or individual emails)
  const loadAndProcessSelectedItems = useCallback(() => {
    setLoading(true);
    setAllLinks([]);

    Office.context.mailbox.getSelectedItemsAsync(async (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        setLoading(false);
        return;
      }

      const selectedItems = asyncResult.value as Office.SelectedItemDetails[];
      if (selectedItems.length === 0) {
        setLoading(false);
        return;
      }
      
      const itemIds = selectedItems.map(item => item.itemId);
      let extractedLinks: string[] = [];

      // Process items sequentially to manage Mailbox resources
      for (const itemId of itemIds) {
        const linksFromItem = await loadAndUnloadSingleItem(itemId); 
        
        linksFromItem.forEach(link => {
            extractedLinks.push(link);
        });
      }

      setAllLinks(extractedLinks);
      setLoading(false);
    });
  }, []); 

  // Office setup and event listener
  useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Outlook) {
        if (Office.context.requirements.isSetSupported("Mailbox", "1.15")) {
          Office.context.mailbox.addHandlerAsync(
            Office.EventType.SelectedItemsChanged,
            loadAndProcessSelectedItems,
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                loadAndProcessSelectedItems();
              }
            }
          );
        } else {
            loadAndProcessSelectedItems();
        }
      }
    });
  }, [loadAndProcessSelectedItems]);


  // Batch Download Logic
  const downloadAllLinks = async () => {
    if (isProcessing) return; 

    setIsProcessing(true);
    setCurrentBatch(0);

    for (let i = 0; i < allLinks.length; i += BATCH_SIZE) {
        const batch = allLinks.slice(i, i + BATCH_SIZE);
        
        batch.forEach((link) => {
            // Open link in a new tab/window for download
            const a = document.createElement("a");
            a.href = link;
            a.target = "_blank";
            a.click();
        });
        
        const linksProcessed = Math.min(i + BATCH_SIZE, allLinks.length);
        setCurrentBatch(linksProcessed);
        
        if (linksProcessed < allLinks.length) {
            // Wait for 10 seconds before opening the next batch
            await new Promise(resolve => setTimeout(resolve, DELAY_MS));
        }
    }

    setIsProcessing(false);
  };

  // --- Attractive UI Rendering ---
  return (
    <Container maxWidth="sm" sx={{ p: 0, pt: 1 }}>
      <Typography
        variant="h5"
        align="center"
        sx={{ 
            fontWeight: 'bold', 
            color: '#106ebe', 
            mb: 2, 
            p: 1,
            borderRadius: 1,
            backgroundColor: '#f5f8ff'
        }}
      >
        Multi-Email Invoice Downloader
      </Typography>

      <Card
        sx={{ 
          maxHeight: 350, 
          overflowY: "auto", 
          mb: 3, 
          boxShadow: 3, 
          borderRadius: 2,
          border: '1px solid #e0e0e0'
        }}
      >
        <CardContent sx={{ p: 0 }}>
          {loading ? (
            <Box display="flex" flexDirection="column" alignItems="center" p={4}>
              <CircularProgress size={30} sx={{ mb: 1, color: '#106ebe' }} />
              <Typography variant="body2" color="text.secondary">
                Analyzing selected messages for files...
              </Typography>
            </Box>
          ) : allLinks.length === 0 ? (
            <Typography variant="body1" color="text.secondary" align="center" p={3}>
              No target files found in selected messages.
            </Typography>
          ) : (
            <List disablePadding>
              {allLinks.map((link, idx) => (
                <React.Fragment key={idx}>
                  <ListItem
                    key={idx} 
                    secondaryAction={
                        <Tooltip title="Open/Download this file">
                            <IconButton 
                                edge="end" 
                                aria-label="open link" 
                                href={link}
                                target="_blank"
                                size="small"
                                sx={{ color: '#106ebe' }}
                                disabled={isProcessing} 
                            >
                                <ArrowDownload16Regular fontSize="small"/>
                            </IconButton>
                        </Tooltip>
                    }
                    sx={{ 
                      px: 2, 
                      py: 0.5,
                      backgroundColor: idx % 2 === 0 ? '#fcfcfc' : 'white', 
                      "&:hover": { backgroundColor: "#e3f2fd" },
                    }}
                  >
                    <ListItemText
                      primary={
                        <Box display="flex" alignItems="center">
                            <Link12Filled style={{ fontSize: '0.9rem', marginRight: 5, color: '#0d47a1' }} />
                            {link}
                        </Box>
                      }
                      primaryTypographyProps={{
                        fontSize: "0.75rem",
                        color: "#555",
                        noWrap: true,
                        sx: {
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          fontFamily: "monospace",
                        },
                      }}
                    />
                  </ListItem>
                  <Divider component="li" light />
                </React.Fragment>
              ))}
            </List>
          )}
        </CardContent>
      </Card>

      {allLinks.length > 0 && (
        <Box textAlign="center">
          <Button
            variant="contained"
            size="large"
            onClick={downloadAllLinks}
            disabled={loading || isProcessing}
            startIcon={isProcessing ? <CircularProgress size={20} color="inherit" /> : <ArrowDownload16Regular />}
            sx={{
              fontWeight: "bold",
              px: 5,
              py: 1.2,
              fontSize: '1rem',
              backgroundColor: isProcessing ? "#ff9800" : "#2e7d32", 
              "&:hover": { backgroundColor: isProcessing ? "#e65100" : "#1b5e20" },
              boxShadow: 4,
            }}
          >
            {isProcessing ? 
                `Opening ${currentBatch}/${allLinks.length} (Next batch in 10s)`
                :
                `Download All Files (${allLinks.length})`
            }
          </Button>
          {(isProcessing && currentBatch === allLinks.length) && (
              <Typography variant="caption" display="block" color="success.main" mt={1}>
                  Process Complete.
              </Typography>
          )}
        </Box>
      )}
    </Container>
  );
};

export default App;