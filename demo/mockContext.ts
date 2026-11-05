import { IDashboardUser } from '../src/components';

/**
 * Stand-in for `this.context.pageContext.user` (plus a job title you'd normally
 * pull from Microsoft Graph `/me`).
 */
export const mockPageContext = {
  user: {
    displayName: 'Alex Morgan',
    email: 'alex.morgan@example.com',
    loginName: 'i:0#.f|membership|alex.morgan@example.com',
  },
};

export const mockUser: IDashboardUser = {
  displayName: mockPageContext.user.displayName,
  email: mockPageContext.user.email,
  jobTitle: 'Product Designer',
};
