'use client';

/**
 * useAuth hook — provides login, logout, token acquisition, and role checking.
 * In demo mode, returns a fake SuperAdmin user without requiring MSAL.
 */
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { useCallback, useMemo } from 'react';
import { loginRequest, apiTokenRequest } from './msal-config';
import type { UserRole } from '@mosy/shared-types';

const isDemoMode = process.env.NEXT_PUBLIC_DEMO_MODE === 'true';

/** Role hierarchy: SuperAdmin > SiteManager > Operator > Viewer. */
const ROLE_HIERARCHY: Record<UserRole, number> = {
  SuperAdmin: 4,
  SiteManager: 3,
  Operator: 2,
  Viewer: 1,
};

export interface AuthUser {
  id: string;
  name: string;
  email: string;
  role: UserRole;
}

const DEMO_USER: AuthUser = {
  id: 'demo-admin-001',
  name: 'Demo Admin',
  email: 'admin@mosy-demo.com',
  role: 'SuperAdmin',
};

function useDemoAuth() {
  const login = useCallback(async () => {}, []);
  const logout = useCallback(async () => {}, []);
  const getAccessToken = useCallback(async () => 'demo-token', []);
  const hasRole = useCallback((_requiredRole: UserRole) => true, []);

  return {
    user: DEMO_USER,
    isAuthenticated: true,
    login,
    logout,
    getAccessToken,
    hasRole,
  };
}

function useMsalAuth() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const user = useMemo((): AuthUser | null => {
    if (!isAuthenticated || accounts.length === 0) return null;
    const account = accounts[0];
    const claims = account.idTokenClaims as Record<string, unknown> | undefined;
    const roles = (claims?.roles as string[]) ?? [];
    const role: UserRole = roles.includes('SuperAdmin')
      ? 'SuperAdmin'
      : roles.includes('SiteManager')
        ? 'SiteManager'
        : roles.includes('Operator')
          ? 'Operator'
          : 'Viewer';

    return {
      id: account.localAccountId,
      name: account.name ?? account.username,
      email: account.username,
      role,
    };
  }, [isAuthenticated, accounts]);

  const login = useCallback(async () => {
    try {
      await instance.loginPopup(loginRequest);
    } catch (error) {
      console.error('[Auth] Login failed:', error);
      throw error;
    }
  }, [instance]);

  const logout = useCallback(async () => {
    try {
      await instance.logoutPopup({ mainWindowRedirectUri: '/auth/signin' });
    } catch (error) {
      console.error('[Auth] Logout failed:', error);
    }
  }, [instance]);

  const getAccessToken = useCallback(async (): Promise<string> => {
    try {
      const result = await instance.acquireTokenSilent({
        ...apiTokenRequest,
        account: accounts[0],
      });
      return result.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const result = await instance.acquireTokenPopup(apiTokenRequest);
        return result.accessToken;
      }
      throw error;
    }
  }, [instance, accounts]);

  /** Check if user has at least the given role level. */
  const hasRole = useCallback(
    (requiredRole: UserRole): boolean => {
      if (!user) return false;
      return ROLE_HIERARCHY[user.role] >= ROLE_HIERARCHY[requiredRole];
    },
    [user]
  );

  return {
    user,
    isAuthenticated,
    login,
    logout,
    getAccessToken,
    hasRole,
  };
}

export function useAuth() {
  if (isDemoMode) {
    return useDemoAuth();
  }
  return useMsalAuth();
}
