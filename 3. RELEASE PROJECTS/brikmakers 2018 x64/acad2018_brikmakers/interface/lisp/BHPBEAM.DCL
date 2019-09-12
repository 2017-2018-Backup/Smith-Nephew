////                  === MD_BEAM.DCL ===
////
////
bhp : dialog {
  label="BHP Steel Sections";
  :row {
    :column {
      :boxed_radio_column { label= "Welded   ";
        :button { label="WB"; key="bhp_wb";}
        :button { label="WC"; key="bhp_wc";}
      }
      spacer;
      :boxed_radio_column { label="Channel";
         :button { label="PFC"; key="bhp_pfc";}
         :button { label="TFC"; key="bhp_tfc";}
         :button { label="TFB"; key="bhp_tfb";}
      }
      spacer;
      :boxed_radio_column { label= "Angles.";
        :button { label="EA"; key="bhp_ea";}
        :button { label="UA"; key="bhp_ua";}
      }
    }
    spacer;
    spacer;
    :column {
      :boxed_radio_column { label= "Universal";
        :button { label="UB"; key="bhp_ub";}
        :button { label="UC"; key="bhp_uc";}
      }
      spacer;
      :boxed_radio_column { label= "Hollow";
        :button { label="CHS"; key="bhp_chs";}
        :button { label="RHS"; key="bhp_rhs";}
        :button { label="SHS"; key="bhp_shs";}
      }
      spacer;
      :boxed_radio_column { label= "Lysaght's";
        :button { label="'C'"; key="bhp_lcp";}
        :button { label="Zed"; key="bhp_lzp";}
      }
    }
  }
  spacer;
  spacer;
  spacer;
  ok_cancel;
}
bhp1_list : dialog {
  label="Member Type";
  :list_box { key="list_box"; tabs= "5"; width= 15; fixed_width= true; height= 10;}
  spacer;
  spacer;
    ok_only;
}
bhp2_list : dialog {
  label="Member Type";
  :list_box { key="list_box"; tabs= "5"; width= 23; fixed_width= true; height= 10;}
  spacer;
  spacer;
    ok_only;
}
det_cen : dialog {
  label = "BHPBEAM Settings";
          : column {
            : radio_row {
              label = "View on Beam";
              width = 10;
              fixed_width = true;
              : radio_button {
                label = "Section";
                mnemonic = "S";
                key = "s_ect";
              }
              : radio_button {
                label = "Plan";
                mnemonic = "P";
                key = "p_lan";
              }
              : radio_button {
                label = "Elevation";
                mnemonic = "E";
                key = "e_lev";
              }
            }
            spacer;
            spacer;
            spacer;
            : row {
              : row {
                : radio_column {
                  label = "Drawing Mode";
                  width = 10;
                  fixed_width = true;
                  : radio_button {
                    label = "Full Detail";
                    mnemonic = "F";
                    key = "f_det";
                  }
                  spacer;
                  : radio_button {
                    label = "Design Mode";
                    mnemonic = "D";
                    key = "d_det";
                  }
                }
              }
              spacer;
              : row {
                : radio_column {
                  label = "Center lines";
                  width = 3;
                  fixed_width = true;
                  : radio_button {
                    label = "Off";
                    mnemonic = "O";
                    key = "cl_of";
                  }
                  spacer;
                  : radio_button {
                    label = "On";
                    mnemonic = "n";
                    key = "cl_on";
                  }
                }
              }
            }
            spacer;
          }
          ok_cancel;
        }
//
bm_ip  : dialog {
         label = "Beam I.P.";
         spacer;
         spacer;
         : boxed_radio_column {
           label = "Section I.P.";
           : radio_button {
             label = "Top Left";
             mnemonic = "L";
             key = "t_opl";
           }
           : radio_button {
             label = "Top Centre";
             mnemonic = "C";
             key = "t_opc";
           }
           : radio_button {
             label = "Top Right";
             mnemonic = "R";
             key = "t_opr";
           }
           spacer;
           spacer;
           spacer;
           : radio_button {
             label = "Middle Left";
             mnemonic = "M";
             key = "m_idl";
           }
           : radio_button {
             label = "Middle Centre";
             mnemonic = "i";
             key = "m_idc";
           }
           : radio_button {
             label = "Middle Right";
             mnemonic = "d";
             key = "m_idr";
           }
           spacer;
           spacer;
           spacer;
           : radio_button {
             label = "Bottom Left";
             mnemonic = "B";
             key = "b_otl";
           }
           : radio_button {
             label = "Bottom Centre";
             mnemonic = "o";
             key = "b_otc";
           }
           : radio_button {
             label = "Bottom Right";
             mnemonic = "t";
             key = "b_otr";
           }
         }
         spacer;
         spacer;
         spacer;
         ok_cancel;
       }
//
mem_loc  : dialog {
         label = "Member Position";
         spacer;
         spacer;
         : boxed_radio_row {
           label = "Beam Location";
           : radio_button {
             label = "Top";
             mnemonic = "T";
             key = "t_op";
           }
           : radio_button {
             label = "Centre";
             mnemonic = "C";
             key = "c_en";
           }
           : radio_button {
             label = "Bottom";
             mnemonic = "B";
             key = "b_ot";
           }
         }
         spacer;
         spacer;
         : boxed_radio_row {
           label = "Beam Location";
           : radio_button {
             label = "Back";
             mnemonic = "B";
             key = "a_ngb";
           }
           : radio_button {
             label = "Gauge";
             mnemonic = "G";
             key = "a_ngg";
           }
           : radio_button {
             label = "Cx/Cy";
             mnemonic = "C";
             key = "a_ngc";
           }
         }
         spacer;
         spacer;
         : radio_row {
           : boxed_radio_column {
             label = "Toes";
             : radio_button {
               label = "Up";
               mnemonic = "U";
               key = "t_oeu";
             }
             : radio_button {
               label = "Down";
               mnemonic = "D";
               key = "t_oed";
             }
           }
           : boxed_radio_column {
             label = "Toes";
             : radio_button {
               label = "Away";
               mnemonic = "A";
               key = "t_oea";
             }
             : radio_button {
               label = "Toward";
               mnemonic = "T";
               key = "t_oet";
             }
           }
           : boxed_radio_column {
             label = "Toes";
             : radio_button {
               label = "Right";
               mnemonic = "R";
               key = "t_oer";
             }
             : radio_button {
               label = "Left";
               mnemonic = "L";
               key = "t_oel";
             }
           }
         }
         spacer;
         spacer;
         : boxed_radio_row {
           label = "Leg View";
           : radio_button {
             label = "Long";
             mnemonic = "L";
             key = "l_egl";
           }
           : radio_button {
             label = "Short";
             mnemonic = "S";
             key = "l_egs";
           }
         }
         spacer;
         spacer;
         ok_cancel;
       }
