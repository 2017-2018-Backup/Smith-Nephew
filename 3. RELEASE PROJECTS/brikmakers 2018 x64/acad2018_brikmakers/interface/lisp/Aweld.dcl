aweld : dialog {
    label = "Weld Selection";
        : row {               
        : column {
        : boxed_radio_column {
            label = "     Weld Side";
            : radio_button {
                label = "This Side";
                key = "this_side";
                value = 1;
            }       
            : radio_button {
                label = "Other Side";
                key = "other_side";
            }       
            : radio_button {
                label = "Both Sides";
                key = "both_side";
            }       
        }
        : boxed_column {
        : toggle {
            label = "Backing Weld";
            key = "ba_weld";    
        }
        : toggle {
            label = "Weld All Around";
            key = "ar_weld";    
        }    
        : toggle {
            label = "Field Weld";
            key = "fi_weld";    
        } 
        }
        }
        spacer_1;
        : column {    
            : edit_box {
                label = "Fillet Size:";
                key = "fill_size";
                allow_accept = false;
            }    
            : edit_box {
                label = "Length-Pitch:";
                key = "pitch_size";
                allow_accept = false;
            }    
            : edit_box {
                label = "Note 1:";
                key = "note1";
                allow_accept = false;
            }    
            : edit_box {
                label = "Note 2:";
                key = "note2";
                allow_accept = false;
            }    
            : edit_box {
                label = "Note 3:";
                key = "note3";
                allow_accept = false;
            }    
    }
}
    spacer;
    ok_cancel;
}    
