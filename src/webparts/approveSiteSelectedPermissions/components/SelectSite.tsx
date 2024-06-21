import * as React from 'react';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import FunctionService from '../../../services/FunctionService';
import styles from './SelectSite.module.scss';
import { ISelectSiteProps } from './ISelectSiteProps';
import { ISite } from '../../../model/ISite';

export const SelectSite: React.FC<ISelectSiteProps> = (props) => {
  const functionService = new FunctionService(props.serviceScope);

  const [newSearchVal, setNewSearchVal] = React.useState<string>("");
  const [sites, setSites] = React.useState<ISite[]>([]);
  const [dropdownOptions, setDropdownOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedSite, setSelectedSite] = React.useState<IDropdownOption>();

  const searchResults = (newValue: string): void => {
    setNewSearchVal(newSearchVal);
    const fetchData = async () => {
      const searchResults = await functionService.searchSites(newValue, 1);
      setSites(searchResults);
    }
  
    // call the function
    fetchData()
      // make sure to catch any error
      .catch(console.error);
  }

  const onPermissionChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    setSelectedSite(item);
    props.siteSelectedCallback(item.key as string);
  };

  React.useEffect(() => {
    const dropdownSiteOptions: IDropdownOption[] = [
      { key: 'permissionssHeader', text: 'Sites', itemType: DropdownMenuItemType.Header }
    ];
    sites.forEach(s => {
      dropdownSiteOptions.push({key: s.Id, text: s.Title, title: s.Url});
    });
    setDropdownOptions(dropdownSiteOptions);
  }, [sites]);

  return (
    <div className={styles.selectSite}>
      <div className={styles.fieldUp}>
        <SearchBox placeholder="Search" onSearch={(newValue) => searchResults(newValue)} />
      </div>
      <div className={styles.fieldDown}>
        <Dropdown
                label="Site selection"
                selectedKey={selectedSite ? selectedSite.key : undefined}
                // eslint-disable-next-line react/jsx-no-bind
                onChange={onPermissionChange}
                placeholder="Select a site"
                options={dropdownOptions} />
      </div>
    </div>
  );
}