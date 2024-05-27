import * as React from 'react';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import styles from './ApproveSiteSelectedPermissions.module.scss';
import * as strings from 'ApproveSiteSelectedPermissionsWebPartStrings';
import type { IApproveSiteSelectedPermissionsProps } from './IApproveSiteSelectedPermissionsProps';
import { SelectSite } from "./SelectSite";
import GraphService from '../../../services/GraphService';

export const ApproveSiteSelectedPermissions: React.FC<IApproveSiteSelectedPermissionsProps> = (props) => {
  const [selectedPermission, setSelectedPermission] = React.useState<IDropdownOption>();
  const [selectedSiteID, setSelectedSiteID] = React.useState<string>(props.siteId);
  const [siteAccess, setSiteAccess] = React.useState<boolean>(false);

  const onPermissionChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    setSelectedPermission(item);
  };

  const dropdownPermissionOptions = [
    { key: 'permissionssHeader', text: strings.PermissionssHeader, itemType: DropdownMenuItemType.Header },
    { key: 'fullcontrol', text: 'Full Control' },
    { key: 'write', text: 'Write' },
    { key: 'read', text: 'Read' }
  ];

  const checkSiteAccess = async () => {
    const graphService = new GraphService(props.serviceScope);
    const isAdmin = await graphService.isSiteAdmin(props.userEMail, props.siteId);
    setSiteAccess(isAdmin);
  };

  const assignPermissions = async () => {
    const graphService = new GraphService(props.serviceScope);
    const appDisplayName = await graphService.servicePrincipal(props.selectedApp);

    const resp = await graphService.grantPermissions(selectedPermission?.key as string, props.selectedApp, appDisplayName, selectedSiteID!);
    console.log(resp);
  };

  const retrieveSiteID = React.useCallback((siteID: string) => {
    setSelectedSiteID(siteID);
    checkSiteAccess();
  }, []);

  React.useEffect(() => {
    checkSiteAccess();
  }, []);
  
  return (
    <section className={`${styles.approveSiteSelectedPermissions} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.field}>          
        <h2>{strings.HeaderLabel}</h2>
      </div>
      
      <div>
        {
          props.isAdminMode?<div className={styles.field}><SelectSite serviceScope={props.serviceScope} siteSelectedCallback={retrieveSiteID} /></div>:<div className={styles.field}><h3>Current site is used</h3></div>
        }
      </div>
      <div className={styles.field}>
        <div className={styles.permDD}>
          <Dropdown
              label={strings.GrantPermissionLabel}
              selectedKey={selectedPermission ? selectedPermission.key : undefined}
              // eslint-disable-next-line react/jsx-no-bind
              onChange={onPermissionChange}
              placeholder={strings.PermissionsPlaceholder}
              options={dropdownPermissionOptions} />
        </div>
      </div>
      <div className={styles.field}>
        <PrimaryButton text={strings.ApprovePrermissionsLabel} onClick={assignPermissions} allowDisabledFocus disabled={!siteAccess} />
      </div>
    </section>
  );
}
