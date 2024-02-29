import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
} from '@microsoft/sp-webpart-base';
import styles from './UserProfileWebpartWebPart.module.scss';
import * as strings from 'UserProfileWebpartWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
 
const mailIcon = require('./assets/icon-mail.png');
const personIcon = require('./assets/icon-person.png');
const schemaIcon = require('./assets/icon-schema.png');
const companyIcon = require('./assets/icon-company.png');
const celebrationIcon = require('./assets/icon-celebration.png');
// // Temporary placeholders for new icons
// const genderIcon = require('./assets/gender-icon.png');
// const userIdIcon = require('./assets/userID-icon.png');
const locationIcon = require('./assets/location-icon.png');
const yearOfJoiningIcon = require('./assets/joining-icon.png');
// const majorIcon = require('./assets/major-icon.png');
// const degreeIcon = require('./assets/degree-icon.png');
// const schoolIcon = require('./assets/school-icon.png');
const defaultImage = require('./assets/default-icon.png');
 
 
export interface IUserProfileWebpartWebPartProps {
  displayBirthday: any;
}
 
interface ListItem {
  UserName: string;
  DateofBirth: string;
  JoiningDate: string;
  // Gender: string;
  Department: string;
  Company: string;
  ProfilePicture: {
    Url: string;
  }
  Email: string;
  RefreshedOn: string;
//properties from the second list
  // School: string
  // Major: string
  // Degree: string
}
 
 
 
 
export default class UserProfileWebpartWebPart extends BaseClientSideWebPart<IUserProfileWebpartWebPartProps> {
 
  private userEmail: string;
  private searchTerm: string = '';
  private usersList: any[] = [];
  suggestionsContainer: HTMLDivElement;
  private usersList1: any[] = [];
  private usersList2: any[] = [];
  private isLoading: boolean = true;
  private lastSearchTerm: string = '';
private lastSearchTimestamp: number = 0;

protected async onInit(): Promise<void> {
  await this.userDetails();
  return super.onInit();
}  

private userPrincipalEmail: string = "";

private async userDetails(): Promise<void> {
  // Ensure that you have access to the SPHttpClient
  const spHttpClient: SPHttpClient = this.context.spHttpClient;

  // Use try-catch to handle errors
  try {
    // Get the current user's information
    const response: SPHttpClientResponse = await spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1);
    const userProperties: any = await response.json();

    console.log("User Details:", userProperties);

    // Access the userPrincipalName from userProperties
    const userPrincipalNameProperty = userProperties.UserProfileProperties.find((property: any) => property.Key === 'SPS-UserPrincipalName');

    if (userPrincipalNameProperty) {
      this.userPrincipalEmail = userPrincipalNameProperty.Value.toLowerCase();
      console.log('User Email using User Principal Name:', this.userPrincipalEmail);
    } else {
      console.error('User Principal Name not found in user properties');
    }
  } catch (error) {
    console.error('Error fetching user properties:', error);
  }
} 

 
 
public getItemsFromSPList(listName: string): Promise<any[]> {
  return new Promise((resolve, reject) => {
    let open = indexedDB.open("MyDatabase", 1);
 
    open.onsuccess = function() {
      console.log("Database opened successfully");
      let db = open.result;
      let tx = db.transaction(listName, "readonly"); // Removed '${}' around listName
      let store = tx.objectStore(listName); // Removed '${}' around listName
 
      let getAllRequest = store.getAll();
 
      getAllRequest.onsuccess = function() {
        resolve(getAllRequest.result);
      };
 
      getAllRequest.onerror = function() {
        reject(getAllRequest.error);
      };
    };
 
    open.onerror = function() {
      reject(open.error);
    };
  });
 
 
}
 
 
 
 
  // private getSharePointListItems(url: string, response: ListItem[]): Promise<ListItem[]> {
  //   return new Promise<ListItem[]>((resolve, reject) => {
  //     this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
  //       headers: {
  //         'Accept': 'application/json;odata=nometadata',
  //         'Content-Type': 'application/json;odata=nometadata',
  //         'odata-version': ''
  //     }
  //     })
  //       .then((spResponse: SPHttpClientResponse) => spResponse.json())
  //       .then((data: any) => {
  //         const items = data.value || []; // Use data.value if it exists, or default to an empty array
 
  //         response = response.concat(items);
 
  //         if (data['odata.nextLink']) {
  //           this.getSharePointListItems(data['odata.nextLink'], response).then(resolve).catch(reject);
  //         } else {
  //           resolve(response);
  //         }
  //       })
  //       .catch((error: any) => {
  //         reject(error);
  //       });
  //   });
  // }
 
  private async getUsersDetails(): Promise<void> {
    try {
 
      this.usersList1 = await this.getItemsFromSPList("SPList");
     
 
 
      console.log("Items retrieved from the first list:", this.usersList1);
     
 
    this.render();
    // Data fetching is complete, update loading state
    this.isLoading = false;
 
    // Trigger a re-render
   
  } catch (error) {
    console.error('Error fetching user details:', error);
  }
}
 
  public async render(): Promise<void> {
    const profileContent: string = this.isLoading ? `
    <div class="${styles["loading"]}">Loading...</div>` :
    `
      <div class="${styles["my-profile"]}">
        <!-- UserProfile will be rendered here -->
      </div>
    `;
 
   
 
    this.domElement.innerHTML = profileContent;
 
    // Get the logged-in user's email
    // this.userEmail = this.context.pageContext.user.email ? this.context.pageContext.user.email.toLowerCase() : '';
    this.renderSearchBar();

    // If it's the initial render, render the profile of the logged-in user
    if (this.isInitialRender) {
        await this.getUsersDetails();
        console.log("Script started");
        this.userEmail = this.context.pageContext.user.email.toLowerCase();
        console.log("Initial User Email:", this.userEmail);
        if(!this.userEmail){
          let extracted = this.userPrincipalEmail.split("#ext#")[0];
          let lastUnderscoreIndex = extracted.lastIndexOf("_");
          if (lastUnderscoreIndex !== -1) {
            extracted = extracted.substring(0, lastUnderscoreIndex) + "@" + extracted.substring(lastUnderscoreIndex + 1);
          }
          this.userEmail = extracted? extracted : '';
    
          console.log("About to log user email");
          console.log("User Email:", this.userEmail);
        if(!this.userEmail){
          return;
        }else{
          this.fetchAndRenderUserProfile(this.userEmail);
        }
        this.isInitialRender = false;
        }else{
        console.log("About to log user email");
        console.log("User Email:", this.userEmail);
        this.fetchAndRenderUserProfile(this.userEmail);
        this.isInitialRender = false;
        }
    } else {
      // If there is a stored search term, render the corresponding profile
      const storedSearchTerm = sessionStorage.getItem('searchTerm');
      if (storedSearchTerm) {
        this.fetchAndRenderUserProfile(storedSearchTerm);
      }
    }
   
   
}
 
protected onDispose(): void {
  // Add any cleanup logic needed
  super.onDispose();
}
 
private isAutomaticRender: boolean = true;
 
 
protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
  // Check if the selectedList property is changed
  if (propertyPath === 'selectedList' && newValue) {
    console.log('Selected List changed. New value:', newValue);
 
    // Check if there's an active search
    const storedSearchTerm = sessionStorage.getItem('searchTerm');
    if (!storedSearchTerm) {
      console.log('Fetching and rendering user profile for logged-in user:', this.userEmail);
      this.fetchAndRenderUserProfile(this.userEmail);
    } else {
      console.log('Fetching and rendering user profile for stored search term:', storedSearchTerm);
      this.fetchAndRenderUserProfile(storedSearchTerm);
    }
  }
 
  super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
}
 
  private searchResults: any[] = [];
 
  private async fetchAndRenderUserProfile(searchTerm?: string): Promise<void> {
    try {
        if (this.usersList1 && this.usersList2) {
            console.log("entering if else for usersList1 and userList2");
            console.log("Search term:", searchTerm);
            let filteredUsersList1 = searchTerm ? this.getFilterBySearchTerm(searchTerm, this.usersList1) : this.usersList1;
            console.log('Filtered users 1:', filteredUsersList1);
           
           
            let allUsers = filteredUsersList1.map(user => {
                // Use each user's Email property instead of the first one
                let userIdFromList1 = user.Title;
                console.log('Searching for user ID for ',user.UserName,":", userIdFromList1);
                // let filteredUsersList2 = userIdFromList1 ? this.getFilterByUserId(userIdFromList1, this.usersList2) : this.usersList2;
                // console.log('Filtered users 2:', filteredUsersList2);
 
                // Create a new object with properties from filteredUsersList1
                let newUser = { ...user };
 
               
 
                // let userFromList2 = filteredUsersList2.find(u => u.Title === user.Title);
 
            // Check if there are multiple items in SPList_Education for the current user
            // if (filteredUsersList2.length > 1) {
            //   // Sort the items by a property (e.g., assuming there is a 'Date' property)
            //   filteredUsersList2.sort((a, b) => new Date(b.Date).getTime() - new Date(a.Date).getTime());
 
            //   // Display the properties from the last item
            //   let lastItem = filteredUsersList2[filteredUsersList2.length - 1];
            //   newUser.Degree = lastItem.Degree;
            //   newUser.School = lastItem.School;
            // } else if (filteredUsersList2.length === 1) {
            //   // If there is only one item, display its properties
            //   newUser.Degree = filteredUsersList2[0].Degree;
            //   newUser.School = filteredUsersList2[0].School;
            // }
 
           
 
            return newUser;
            });
 
            // Render all users
            allUsers.forEach(user => {
                console.log("entering if else for allUsers");
                console.log('User details:', user);
                this.renderUserProfile(user);
            });
        }
    } catch (error) {
        console.error('Error fetching data:', error);
    }
}
 
 
  private clearUserProfile(): void {
    const profileElement = this.domElement.querySelector(`.${styles['my-profile']}`);
    if (profileElement) {
      profileElement.innerHTML = ''; // Clear the profile content
    }
  }
 
  private adjustDateForTimeZone(dateString: string) {
    // Add your timezone adjustment logic here
    const timeZoneDifferenceHours = 5; // Adjust this based on your timezone
    const timeZoneDifferenceMinutes = 30;
 
    const date = new Date(dateString);
    date.setHours(date.getHours() + timeZoneDifferenceHours);
    date.setMinutes(date.getMinutes() + timeZoneDifferenceMinutes);
 
    return date;
  }
 
  private renderUserProfile(userDetails: any): void {
    console.log('Rendering user profile for:', userDetails.UserName);
    console.log('User details before rendering:', userDetails);
 
    const dateofbirth = userDetails.DateofBirth ? this.adjustDateForTimeZone(userDetails.DateofBirth).toISOString().substring(5, 10) : 'Nil';
    const dateofbirthformatted = userDetails.DateofBirth ? this.formatBirthday(dateofbirth): 'Nil';
 
    const joiningdate = userDetails.JoiningDate ? this.adjustDateForTimeZone(userDetails.JoiningDate).toISOString().substring(0, 10) : 'Nil';
    const joiningdateformatted = userDetails.JoiningDate ? this.formatDate(joiningdate) : 'Nil';  
 
    const profileContent = `
      <h3>My Profile</h3>
      <div class="${styles["top-section"]}">
<img src="${userDetails.ProfilePicture && userDetails.ProfilePicture.Url ? userDetails.ProfilePicture.Url : defaultImage}" alt="Profile Picture" class="${styles["profile-dp"]}"onError="this.onerror=null;this.src='${defaultImage}';">
<div class="${styles["profile-content"]}">
 
         
<h4>${userDetails.UserName || 'Nil'}</h4>
 
            <h6>
            <span>${userDetails.JobTitle || 'Nil'}</span>
            <span>${userDetails.Department || 'Nil'}</span>
          </h6>
        </div>
      </div>
 
      <div class="${styles["bottom-section"]}">
        <h4>Contact information</h4>
        <div class="${styles["inner-contents"]}">
 
        <a>
        <img src="${companyIcon}">
        <span>Group Company<br> <i id = "Group company">${userDetails.GroupCompany}</i></span>
      </a>
 
          <a>
            <img src="${mailIcon}">
            <span>Email <br> <i id = "userEmail">${userDetails.Email}</i></span>
          </a>
 
 
          <a>
            <img src="${personIcon}">
            <span>Job title <br> <i class="${styles["color-black"]}">${userDetails.JobTitle}</i></span>
          </a>
 
          <a>
            <img src="${schemaIcon}">
            <span>Department <br> <i class="${styles["color-black"]}">${userDetails.Department}</i></span>
          </a>
 
         
 
          <a>
            <img src="${celebrationIcon}">
            <span>Date Of Birth (dd-mm) <br> <i class="${styles["color-black"]}">${dateofbirthformatted}</i></span>
          </a>
 
          <!-- New columns with temporary placeholders for icons -->
 
          <a>
          <img src="${locationIcon}">
            <span> Location <br> <i class="${styles["color-black"]}">${userDetails.Location || 'Nil'}</i></span>
          </a>
 
          <a>
          <img src="${yearOfJoiningIcon}">
            <span>Date of Joining <br> <i class="${styles["color-black"]}">${joiningdateformatted || 'Nil'}</i></span>
          </a>
 
        </div>
      </div>
    `;
 
    (this.domElement.querySelector(`.${styles["my-profile"]}`) as HTMLElement).innerHTML = profileContent;
 
    // Get the DOM element where the user profile should be displayed
const userProfileElement = this.domElement.querySelector(`.${styles["my-profile"]}`) as HTMLElement;
 
// Check if the DOM element exists
if (userProfileElement) {
    // Insert the user profile content into the DOM element
    userProfileElement.innerHTML = profileContent;
} else {
    console.error('User profile element not found.');
}
 
    const emailElement = this.domElement.querySelector('#userEmail');
    if (emailElement) {
        emailElement.addEventListener('click', () => {
            // Open the default email client with a new email
            window.location.href = `mailto:${userDetails.Email}`;
        });
    } else {
        console.error('Email element not found in the DOM.');
        console.log('User profile rendered successfully.');
    }
  }
 
  private formatDate(date: string): string {
    const [year, month, day] = date.split('-');
    return `${day}-${month}-${year}`;
  }
 
  private formatBirthday(date: string): string {
    const [month, day] = date.split('-');
    return `${day}-${month}`;
  }
 
 
  private getFilterBySearchTerm(searchTerm: string, usersList: any[]): any[] {
    if (typeof searchTerm !== 'string') {
      searchTerm = '';
    }
 
    if (!searchTerm.trim()) {
      // If no valid search term is provided, return the original array
      return usersList;
    }
 
    const searchableColumns = ['UserName', 'Email', 'Title'];
    const searchTerms = searchTerm.toLowerCase().split(' '); // Split search term into individual words
    const filteredUsers = usersList.filter(user => {
      // Check if any of the search terms is a substring of any user property in the searchable columns
      return searchableColumns.some(column => {
        const propertyValue = user[column];
 
        // Check if the property value is not null or undefined before performing operations on it
        return (
          propertyValue &&
          typeof propertyValue === 'string' &&
          searchTerms.every(searchWord =>
            propertyValue.toLowerCase().includes(searchWord)
          )
        );
      });
    });
 
    return filteredUsers;
  }
 
 
 
  private getFilterByUserId(searchTerm: string, usersList: any[]): any[] {
    if (!searchTerm) {
      return usersList;
    }
 
    // const searchTermLowerCase = searchTerm.toLowerCase();
    const filteredUsers = usersList.filter(user => {
      const userUserId = user['Title']?.trim().toLowerCase();
      const searchTermLowerCase = searchTerm.trim().toLowerCase();
      return userUserId === searchTermLowerCase;
 
    });
 
    return filteredUsers;
  }
   
 
  private isInitialRender: boolean = true;
 
  private renderSearchBar(): void {
    const profileElement = this.domElement.querySelector(`.${styles["my-profile"]}`);
    if (!profileElement) {
      return;
    }
 
    const searchBarContent = `
      <div class="${styles["search-bar"]}">
        <input type="text" id="searchInput" placeholder="Search for People Across RPG..." />
        <div class="${styles["suggestions"]}"></div>
        <button id="searchButton">Search</button>
      </div>
    `;
 
    profileElement.insertAdjacentHTML('beforebegin', searchBarContent);
 
    const searchInput = document.getElementById('searchInput') as HTMLInputElement;
    const searchButton = document.getElementById('searchButton');
    const suggestionsContainer = document.querySelector(`.${styles["suggestions"]}`) as HTMLDivElement;
 
    if (!searchInput || !searchButton || !suggestionsContainer) {
      return;
    }
 
    // Attach event listeners
const handleSearchButtonClick = (event: MouseEvent) => {
  event.stopPropagation(); // Prevent the event from bubbling up to the document
  console.log('Search button clicked.');
  const searchTerm = searchInput.value.trim();
  console.log('Search term:', searchTerm);
  suggestionsContainer.innerHTML = ''; // Clear suggestions
  this.fetchAndRenderSuggestions(searchTerm);
};
 
const handleInputKeyDown = (event: KeyboardEvent) => {
  if (event.key === 'Enter') {
    console.log('Enter key pressed.');
    handleSearchButtonClick(new MouseEvent('click'));
  }
};
 
searchButton.addEventListener('click', handleSearchButtonClick);
searchInput.addEventListener('keydown', handleInputKeyDown);
 
// Add event listener to document to handle clicks outside the suggestions list
document.addEventListener('click', () => {
  // Click is outside the suggestions list, clear suggestions
  suggestionsContainer.innerHTML = '';
});
  }
 
 
private handleSuggestionClick(searchTerm: string): void {
  console.log('Inside handleSuggestionClick, called with:', searchTerm);
 
  // Log the entire usersList1 array
  console.log('All users in usersList1:', this.usersList1);
 
 
  // Find the selected user in usersList1
  const selectedUser = this.usersList1.find(user => {
    user.Email = user.Email ?? 'Nil';
    const isMatch = user.Email.toLowerCase() === searchTerm.toLowerCase();
    if (isMatch) {
      console.log('Match found for:', user.UserName);
    }
    return isMatch;
  });
 
  if (selectedUser) {
    // Log the selected user
    console.log('Selected user:', selectedUser);
 
    // Clear existing profile and render the selected user
    this.clearUserProfile();
    this.renderUserProfile(selectedUser);
 
    // Clear suggestions after rendering the profile
    const suggestionsContainer = document.querySelector(`.${styles["suggestions"]}`) as HTMLDivElement;
    if (suggestionsContainer) {
      suggestionsContainer.innerHTML = '';
    }
 
  } else {
    console.log('No user found for search term:', searchTerm);
  }
}
 
 
// private handleSuggestionClick(searchTerm: string): void {
//   this.fetchAndRenderUserProfile(searchTerm);
 
//   // Store the search term in sessionStorage
//   sessionStorage.setItem('searchTerm', searchTerm);
// }
 
 
 
// private async fetchAndRenderSuggestionsDelayed(searchTerm: string, container: HTMLDivElement): Promise<void> {
//   try {
//     if (this.usersList1) {
//       this.fetchAndRenderSuggestionsDelayed(searchTerm, container);
//       // Update the last search term and timestamp
//       this.lastSearchTerm = searchTerm;
//       this.lastSearchTimestamp = new Date().getTime();
//     } else {
//       console.error('User list is empty.');
//     }
//   } catch (error) {
//     console.error('Error fetching data:', error);
//   }
// }
 
private fetchAndRenderSuggestions(searchTerm: string): void {
  const searchInput = document.getElementById('searchInput') as HTMLInputElement;
  const suggestionsContainer = document.querySelector(`.${styles["suggestions"]}`) as HTMLDivElement;
 
  if (searchInput && suggestionsContainer) {
    console.log("search input:", searchInput.value);
    const trimmedSearchTerm = searchTerm.trim();
    console.log('Rendering suggestions for searchTerm:', trimmedSearchTerm);
 
    // Perform the logic to fetch and render suggestions based on the searchTerm
    const filteredSuggestions = this.getFilterBySearchTerm(trimmedSearchTerm, this.usersList1);
    console.log('Filtered suggestions:', filteredSuggestions);
 
    this.renderSuggestions(filteredSuggestions, suggestionsContainer);
  }
}
 
 
private searchDelayTimeout: number;
 
 
 
  private renderSuggestions(suggestions: any[], container: HTMLDivElement): void {
    container.innerHTML = ''; // Clear previous suggestions
 
    if (suggestions.length > 0) {
        // Sort suggestions by UserName in alphabetical order
        // suggestions.sort((a, b) => a.UserName.localeCompare(b.UserName));
 
        const suggestionList = document.createElement('ul');
        suggestionList.classList.add(styles['suggestion-list']);
 
        suggestions.forEach((suggestion) => {
            const suggestionItem = document.createElement('li');
            suggestionItem.classList.add(styles['suggestion-item']);
 
            const UserName = suggestion.UserName ?? 'Nil';
            const jobTitle = suggestion.JobTitle ?? 'Nil';
            const Email = suggestion.Email ?? 'Nil';
 
            suggestionItem.innerHTML = `
                <span class="${styles['suggestion-item-title']}">${UserName}</span>
                <span class="${styles['suggestion-dot']}">·</span>
                <span class="${styles['suggestion-item-jobtitle']}">${jobTitle}</span>
                <span class="${styles['suggestion-dot']}">·</span>
                <span class="${styles['suggestion-item-email']}">${Email}</span>`;
 
            suggestionItem.addEventListener('click', () => {
                // Handle suggestion click
                const titleElement = suggestionItem.querySelector(`.${styles['suggestion-item-email']}`);
                const title = titleElement?.textContent?.trim() || '';
                const UserID = suggestion.Title ?? 'Nil';
                this.fetchAndRenderUserProfile(UserID);
 
                 // Clear suggestions
    container.innerHTML = '';
            });
 
            suggestionList.appendChild(suggestionItem);
        });
 
        container.appendChild(suggestionList);
    } else {
        // If no suggestions, you can display a message or hide the container
        container.textContent = 'No suggestions found.';
    }
 
    // Add event listener to document to handle clicks outside the suggestions list
document.addEventListener('click', (event) => {
  const target = event.target as Node;
  if (!container.contains(target)) {
    // Click is outside the suggestions list, clear suggestions
    container.innerHTML = '';
  }
});
}
 
 
 
protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: {
          description: strings.PropertyPaneDescription,
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
              // PropertyPaneDropdown components removed from here
            ],
          },
        ],
      },
    ],
  };
}
}