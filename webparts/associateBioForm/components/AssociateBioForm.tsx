// import "core-js";
// import '@pnp/polyfill-ie11';


// import ProxyPolyfillBuilder from 'proxy-polyfill/src/proxy';
// import 'react-app-polyfill/ie11';
import * as React from 'react';
import styles from './AssociateBioForm.module.scss';
import { useEffect, useState } from 'react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp/presets/all";
import '@pnp/sp/presets/all';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { TextField, Button } from '@material-ui/core';
import Snackbar from '@material-ui/core/Snackbar';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import MuiAlert, { AlertProps } from '@material-ui/lab/Alert';
import './CommonStylesheet.scss';




const AssociateBioForm = (props) => {
  // const proxyPolyfill = ProxyPolyfillBuilder();
  const [defaultUser, setDefaultUser] = useState('');
  const [userId, setUserId] = useState();
  const [disableSubmit, setDisableSubmit] = useState(true);
  const [bio, setBio] = useState([]);
  const [error, setError] = useState(false);
  const [charLeft, setCharLeft] = useState(1100);
  const [listData, setListData] = useState([]);
  const [disable, setDisable] = useState(false);
  const [spinner, showSpinner] = useState(false);
  const [open, setOpen] = useState(false);

  const CHARACTER_LIMIT = 1100;
  let userInfo: any;
  let filterCond: any;
  let userAccess: any;
  let validity: boolean;

  const Alert = function (props: AlertProps) {
    return <MuiAlert elevation={6} variant="filled" {...props} />;
  };
  const getListData = function (cond) {

    sp.web.lists.getByTitle("Associate Bios").items.filter(cond).get().then(function (listItem: any) {
      console.log(listItem);
      if (listItem.length > 0) {
        console.log(listItem);
        setListData(listItem[0]);
        setBio(listItem[0]['Bio']);
      }

    }, (err)=> {
      console.log("err", err);
    }).catch(function (err) { console.log(err) });

    // props.context.spHttpClient.get(props.context.pageContext.web.absoluteUrl + "_api/lists/GetByTitle('Associate Bios')/items?$filter=" + cond,
    //   SPHttpClient.configurations.v1)
    //   .then((response: SPHttpClientResponse) => {
    //     response.json().then((responseJSON: any) => {
    //       console.log(responseJSON);
    //     });
    //   }).then((listItem: any) => {
    //     if (listItem.length > 0) {
    //       console.log(listItem);
    //       setListData(listItem[0]);
    //       setBio(listItem[0]['Bio']);
    //     }
    //   }, function (err) {
    //     console.log("err", err);
    //   }).catch(function (err) { console.log(err) });
  };



  const getPeoplePickerItems = function (items: any[]) {
    if (items.length > 0) {
      userInfo = items[0];
      setUserId(userInfo.id);
      filterCond = "UserId eq " + userInfo.id;
      setDefaultUser(userInfo.secondaryText);
      getListData(filterCond);
    }
    else {
      setBio([]);
    }
  };


  const getUserDetail = function () {
    if (props) {
      setDefaultUser(props.context.pageContext.user.email);
      filterCond = "Title eq " + props.context.pageContext.user.displayName;
      sp.web.siteUsers.getByEmail(props.context.pageContext.user.email).get().then(function (response) {
        getListData(filterCond);
        setUserId(response.Id);
        setDisableSubmit(true);
      }, (err)=> {
        console.log("err", err);
      }).catch(function (err) { console.log(err) });
    }
  };

  const checkValidity = function (event) {
    if (!event.target.value) {
      validity = false;
    }
    else {
      validity = true;
    }
    return validity;
  };

  const handleChange = function (event) {
    if (checkValidity(event)) {
      const charCount = event.target.value.length;
      const charLeft = CHARACTER_LIMIT - charCount;
      setCharLeft(charLeft);
      setError(false);
      setBio(event.target.value);
      setDisableSubmit(false);
    } else {
      setError(true);
      setDisableSubmit(true);
      setBio(event.target.value);
    }
  };
  const checkGoupPermissions = function () {
    const url = "https://edifecs.sharepoint.com/sites/PC/_api/web/sitegroups/getbyname('People & Culture Owners')/CanCurrentUserViewMembership";
    props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(function (response: SPHttpClientResponse): Promise<any> {
        return response.json();
      }).then(function (userPemission) {
        userAccess = !userPemission.value;
        setDisable(userAccess);
      }, (err)=> {
        console.log("err", err);
      }).catch(function (err) { console.log(err) });
  };

  useEffect(function () {
    getUserDetail();
    checkGoupPermissions();
  }, []);

  const handleClose = function (event?: React.SyntheticEvent, reason?: string) {
    if (reason === 'clickaway') {
      return;
    }
    setOpen(false);
    setTimeout(function () { window.location.reload(); }, 2000);
  };

  const handleSubmit = function (e) {
    e.preventDefault();
    showSpinner(true);
    if (userId == listData['UserId']) {
      sp.web.lists.getByTitle("Associate Bios").items.getById(listData['Id']).update({ 'Bio': bio }).then(function (items) {
        if (items) {
          setOpen(true);
          showSpinner(false);
          setTimeout(function () { window.location.reload(); }, 2000);
        }
      }, (err)=> {
        console.log("err", err);
      }).catch(function (err) { console.log(err) });

    } else {
      sp.web.lists.getByTitle("Associate Bios").items.add({ 'Bio': bio, 'UserId': userId }).then(function (response) {
        console.log(response);
        if (response) {
          showSpinner(false);
          setOpen(true);
          setTimeout(function () { window.location.reload(); }, 2000);
        }
      }, (err) =>{
        console.log("err", err);
      }).catch(function (err) { console.log(err) });
    }
  };

  return (
    <div className="associateBioFormWrapper">

      {spinner ? <div className="msSpinner">
        <Spinner label="Submitting data, wait..." size={SpinnerSize.large} />
      </div> : open ? <Snackbar open={open} autoHideDuration={5000} onClose={handleClose} style={{ height: "100%" }}>
        <Alert onClose={handleClose} severity="success">
          Thank you for filling out your information!
               </Alert>
      </Snackbar> :
           <form onSubmit={handleSubmit} className={disable ? "disabled" : ""}>
              <PeoplePicker
                context={props.context} ensureUser={true} defaultSelectedUsers={[defaultUser]} titleText="Associate Name" personSelectionLimit={1} showtooltip={true} isRequired={true} disabled={disable} selectedItems={getPeoplePickerItems} showHiddenInUI={false} principalTypes={[PrincipalType.User]} resolveDelay={500} />
              <span className="message">"You can create or update Bios for other associates by changing the user above.  If they've already submitted a Bio, you will be able to make changes or enter a new one if they have not yet."</span>
              <TextField id="outlined-textarea" className="MuiFormControl-root MuiTextField-root MuiFormControl-marginNormal MuiFormControl-fullWidth" required label="Bio" inputProps={{ maxlength: CHARACTER_LIMIT }} name="Bio" value={bio} helperText={'Characters Limit ' + charLeft} onChange={handleChange} margin="normal" variant="outlined" placeholder="Enter Bio here..." multiline onBlur={handleChange} />
              {error && <p className="errorText">This field is required</p>}
              <Button variant="contained" color="primary" disabled={disableSubmit} type="submit">Submit</Button>
            </form>
      }
    </div>
  );
};
export default AssociateBioForm;