/* eslint-disable no-useless-catch */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import "@pnp/sp/files";
import "@pnp/sp/site-users/web";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files/web";

import {
  // TextField,
  PrimaryButton,
  DefaultButton,
 
  Modal,
  IconButton,
  Icon,
} from "@fluentui/react";
import { mergeStyleSets } from "@fluentui/react/lib/Styling";
import CryptoJS from "crypto-js";

export interface IPasscodeModalProps {
  sp: any;
  user: any;
  isOpen: boolean;
  onClose: () => void;
  onSuccess: () => void;
  createPasscodeUrl: string; // Added property for redirect URL
  _makeIsPassCodeValidateFalse: any;
}

export interface IPasscodeModalState {
  userId: any;
  passcode: string;
  errorMessage: string;
  userPasscodes: Array<{ username: string; passcode: string }>;
  userEmail: string;
  isCreating: boolean;
  isPasswordVisible: boolean;
}

export default class PasscodeModal extends React.Component<
  IPasscodeModalProps,
  IPasscodeModalState
> {
  // private encryptionKey: string = 'default_secret_key'; // Use a secure key in production
  private key = CryptoJS.enc.Utf8.parse("b75524255a7f54d2726a951bb39204df");
  private iv = CryptoJS.enc.Utf8.parse("1583288699248111");
  constructor(props: IPasscodeModalProps) {
    super(props);

    this.state = {
      passcode: "",
      errorMessage: "",
      userPasscodes: [],
      userEmail: this.props.user.email,
      userId: "",
      isCreating: false,
      isPasswordVisible: false,
    };
  }

  public async componentDidMount() {
    // console.log("Component did mount");
    await this.fetchStoredPasscodes();
    const userId = await this.getUserIdByEmail(this.props.user.email);
    this.setState({ userId });
    // console.log("User ID:", userId);
  }

  private getUserIdByEmail = async (email: string): Promise<number> => {
    try {
      const user = await this.props.sp.web.siteUsers.getByEmail(email)();
      // console.log("Fetched user:", user);
      return user.Id;
    } catch (error) {
      // console.error("Error fetching user ID:", error);
      throw error;
    }
  };

  private fetchStoredPasscodes = async () => {
    const user = await this.props.sp?.web.currentUser();
    // console.log("Current user:", user);

    try {
      const items: any[] = await this.props.sp.web.lists
        .getByTitle("passcodes")
        .items.filter(`UserId eq ${user.Id}`)
        .select("User/EMail", "User/Title", "passcode")
        .expand("User")();

      const userPasscodes = items.map((item) => {
        const decryptedPasscode = this.decrypt(item.passcode);
        return {
          username: item.User.Title,
          passcode: decryptedPasscode,
        };
      });

      this.setState({ userPasscodes }, this.checkUserPasscode);
      // console.log("Fetched passcodes:", userPasscodes);
    } catch (error) {
      console.error("Error fetching passcodes:", error);
      this.setState({ errorMessage: "Failed to fetch passcodes." });
    }
  };

  private checkUserPasscode = () => {
    const { userPasscodes } = this.state;
    const userPasscode = userPasscodes.find(
      (up) => up.username === this.props.user.displayName
    );

    if (!userPasscode) {
      this.setState({ isCreating: true });
    }
    // console.log("User passcode check:", userPasscode);
  };

  private onPasscodeChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ passcode: event.target.value });
  };

  // private togglePasswordVisibility = () => {
  //   this.setState({ isPasswordVisible: !this.state.isPasswordVisible });
  //   // console.log("Password visibility toggled:", this.state.isPasswordVisible);
  // };

  private decrypt = (encryptedText: string): string => {
    const bytes = CryptoJS.AES.decrypt(encryptedText, this.key, {
      iv: this.iv,
    });
    const decrypted = bytes.toString(CryptoJS.enc.Utf8);
    return decrypted;
  };

  private validatePasscode = () => {
    const { passcode, userPasscodes } = this.state;
    const userPasscode = userPasscodes.find(
      (up) => up.username === this.props.user.displayName
    );

    if (!userPasscode) {
      this.setState({ errorMessage: "No passcode found for this user." });
      // console.log("No passcode found for user");
      return;
    }

    if (userPasscode.passcode === passcode) {
      this.props.onSuccess();
      this.props.onClose();
      this.setState({ passcode: "" });
      // console.log("Passcode validated successfully");
    } else {
      this.setState({ errorMessage: "Invalid passcode. Please try again." });
      // console.log("Invalid passcode");
    }
  };

  private redirectToCreatePasscode = () => {
    // console.log("Redirecting to create passcode URL:", this.props.createPasscodeUrl);
    window.location.href = this.props.createPasscodeUrl;
  };

  private togglePasscodeVisibility = () => {
    this.setState((prevState) => ({
      isPasswordVisible: !prevState.isPasswordVisible,
    }));
    setTimeout(() => {
      this.setState({ isPasswordVisible: false });
    }, 500);
  };

  public render(): React.ReactElement<IPasscodeModalProps> {
    const { isOpen, onClose } = this.props;
    const {
      passcode,
      errorMessage,
      isCreating,
      // isPasswordVisible,
    } = this.state;

    const styles = mergeStyleSets({
      modal: {
        padding: "10px",
        minWidth: "300px",
        maxWidth: "80vw",
        width: "100%",
        "@media (min-width: 768px)": {
          maxWidth: "580px",
        },
        "@media (max-width: 767px)": {
          maxWidth: "290px",
        },
        margin: "auto",
        backgroundColor: "white",
        borderRadius: "4px",
        boxShadow: "0 2px 8px rgba(0, 0, 0, 0.26)",
      },
      header: {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        borderBottom: "1px solid #ddd",
        marginTop: "4px",
        marginBottom: "4px",
        fontWeight: "400",
        padding: "5px",
      },
      headerTitle: {
        margin: "5px",
        marginLeft: "5px",
        fontSize: "16px",
        fontWeight: "400",
      },
      body: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        textAlign: "center",
        padding: "20px 0",
        height: "120px",
      },
      contentContainer: {
        width: "70%",
        display: "flex",
        flexDirection: "column",
        justifyContent: "center",
        alignItems: "center",
      },
      // footer: {
      //   display: "flex",
      //   justifyContent: "end",
      //   marginTop: "20px",
      //   borderTop: "1px solid #ddd",
      //   paddingTop: "10px",
      // },
      // button: {
      //   margin: "0 5px",
      // },
      iconButton: {
        marginRight: "10px",
      },
      buttonText: {
        fontSize: "16px",
        fontWeight: "bold",
        color: "#0078d4",
      },
      errorMessage: {
        color: "red",
        marginTop: "10px",
        alignSelf: "center",
      },
      noHover: {
        ":hover": {
          transform: "none !important",
          transition: "none !important",
          boxShadow: "none !important",
          backgroundColor: "transparent !important",
        },
      },

      footer: {
        display: "flex",
        justifyContent: "space-between", // Adjusted to space between
        // marginTop: '20px',
        borderTop: "1px solid #ddd",
        paddingTop: "10px",
      },
      button: {
        flex: "1 1 50%", // Ensures each button takes up 50% of the footer width
        margin: "0 5px", // Adds some space between the buttons
      },
      buttonContent: {
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      },
      buttonIcon: {
        marginRight: "4px", // Adjust the space between the icon and text
      },

      removeTopMargin: {
        marginTop: "4px",
        marginBottom: "4px",
      },
    });

    // function _makeIsPassCodeValidateFalse() {
    //   throw new Error("Function not implemented.");
    // }

    return (
      <Modal
        isOpen={isOpen}
        onDismiss={onClose}
        isBlocking={true}
        containerClassName={styles.modal}
      >
        <div className={styles.header}>
          <div style={{ display: "flex", alignItems: "center" }}>
            <IconButton iconProps={{ iconName: "Lock" }} />
            <h4 className={styles.headerTitle}>Passcode Verification</h4>
          </div>
          <IconButton iconProps={{ iconName: "Cancel" }} onClick={onClose} />
        </div>
        <div className={styles.body} style={{ textAlign: "center" }}>
          {isCreating ? (
            <>
              <p>
                Passcode is not set. Please create a passcode to proceed
                further.
              </p>
             
            </>
          ) : (
            <>
              <div className={styles.contentContainer}>
                <label>Enter your passcode for verification:</label>
                {/* <TextField
                  value={passcode}
                  canRevealPassword
                  onChange={this.onPasscodeChange}
                  type="password"
                  styles={{ root: { width: "100%", marginTop: "5px" } }} // This line ensures the TextField occupies the full width
                /> */}
                <div style={{ width: "100%", marginTop: "5px", position: "relative" }}>
  <input
    type={this.state.isPasswordVisible ? "text" : "password"}
    value={passcode}
    onChange={this.onPasscodeChange}
    maxLength={6}
    pattern="\d*"
    title="Please enter a 6-character combination of alphabets and numbers"
    style={{
      width: "100%",
      padding: "8px 40px 8px 8px", // Extra padding for space for the icon
      boxSizing: "border-box",
      border: "1px solid #ccc",
      borderRadius: "4px",
      fontSize: "14px",
    }}
  />
  <button
    type="button"
    onClick={this.togglePasscodeVisibility}
    style={{
      position: "absolute",
      right: "10px",
      top: "50%",
      transform: "translateY(-50%)",
      backgroundColor: "transparent",
      border: "none",
      cursor: "pointer",
      padding: 0,
    }}
    aria-label={this.state.isPasswordVisible ? "Hide passcode" : "Show passcode"}
  >
    <Icon
      iconName={this.state.isPasswordVisible ? "View" : "Hide"}
      style={{ fontSize: "18px", color: "#666" }}
    />
  </button>
</div>
                {errorMessage && (
                  <span className={styles.errorMessage}>{errorMessage}</span>
                )}
              </div>
            </>
          )}
        </div>
        <div className={styles.footer}>
          {isCreating?
          <PrimaryButton
          className={styles.button}
          text="Create Passcode"
          styles={{ root: styles.buttonContent }}
          onClick={this.redirectToCreatePasscode}
          iconProps={{ iconName: "CheckedOutByOther12" }}
        />:<PrimaryButton
            className={styles.button}
            text="Verify"
            styles={{ root: styles.buttonContent }}
            iconProps={{ iconName: "CheckedOutByOther12" }}
            onClick={this.validatePasscode}
          />}
          <DefaultButton
            className={styles.button}
            text="Cancel"
            onClick={() => {
              onClose();
              this.setState({ passcode: "" });
            }}
            styles={{ root: styles.buttonContent }}
            iconProps={{ iconName: "ErrorBadge" }}
          />
        </div>
      </Modal>
    );
  }
}
