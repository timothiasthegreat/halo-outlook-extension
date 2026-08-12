# Azure AD App Registration

The watcher uses Microsoft Graph API to read messages from Exchange. This requires an Azure AD app registration with the right permissions.

---

## Step 1: Create the App Registration

1. Go to **[Azure Portal → App Registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)**
2. Click **New Registration**
3. Fill in:
   - **Name:** `Halo Outlook Watcher`
   - **Supported account types:** Accounts in this organizational directory only (single tenant)
   - **Redirect URI:** Leave blank (the watcher uses client credentials, no redirect needed)
4. Click **Register**

5. Copy the **Application (client) ID** — this is `graph.client_id`
6. Copy the **Directory (tenant) ID** — this is `graph.tenant_id`

```yaml
graph:
  tenant_id: "a1b2c3d4-..."   # Directory (tenant) ID
  client_id: "d4c3b2a1-..."   # Application (client) ID
```

---

## Step 2: Create a Client Secret

1. In your new app registration, go to **Certificates & Secrets**
2. Click **New Client Secret**
3. Fill in:
   - **Description:** `Watcher access`
   - **Expires:** 24 months (or your org's policy)
4. Click **Add**
5. **Copy the Value immediately** — you won't see it again

```yaml
graph:
  client_secret: "the-client-secret-value"
```

---

## Step 3: Add API Permissions

The watcher needs `Mail.Read` application permission to read messages from the configured mailbox.

1. Go to **API Permissions**
2. Click **Add a permission**
3. Choose **Microsoft Graph**
4. Choose **Application permissions** (NOT delegated — the watcher runs unattended)
5. Search for `Mail.Read`
6. Check **Mail.Read** (Read mail in all mailboxes — or `Mail.ReadBasic.All` for metadata only)
7. Click **Add permissions**

### Permission details:

| Permission | What it grants | When to use |
|---|---|---|
| `Mail.Read` | Read all mail in the configured mailbox | Standard — needed for conversationId filtering |
| `Mail.ReadBasic.All` | Read mail metadata only (subject, from, date) | If you only need metadata, not body content |
| `User.Read.All` | Read user profile info | Optional — for user lookup by email |

The watcher uses `GET /users/{user_email}/messages` which requires `Mail.Read`.

---

## Step 4: Grant Admin Consent

Application permissions require admin consent.

1. Go to **API Permissions**
2. Click **Grant admin consent for [tenant name]**
3. Confirm

The permissions table should show **Admin consent required: Yes** with a green checkmark.

---

## Step 5: Verify

Run the setup check to confirm Graph connectivity:

```bash
python scripts/setup_check.py
```

Expected output:

```
✓ Graph API reachable → https://graph.microsoft.com/v1.0
✓ Token acquired (scope: https://graph.microsoft.com/.default)
✓ Mailbox accessible: tim@firesideit.ca
```

---

## Troubleshooting

| Error | Likely cause | Fix |
|---|---|---|
| `AADSTS700016: Application not found` | Wrong client_id or tenant_id | Verify both IDs in Azure Portal |
| `AADSTS7000215: Invalid client secret` | Expired or wrong secret | Create a new client secret |
| `AADSTS65001: Consent required` | Admin consent not granted | Go to API Permissions → Grant admin consent |
| `ErrorAccessDenied` from Graph | Missing Mail.Read permission | Verify permission is Application type (not Delegated), admin consent granted |
| `ResourceNotFound` for `/users/{email}` | User email doesn't match a mailbox | Verify the email address — it must match a licensed Exchange Online mailbox |

---

## Least-Privilege Alternative

If your security policy requires least privilege, you can restrict the app to a single mailbox using **Application Access Policy**:

1. Grant `Mail.Read` application permission as above
2. Create a mail-enabled security group containing only the watched mailbox
3. Run: `New-ApplicationAccessPolicy -AppId <client_id> -PolicyScopeGroupId <group_email> -AccessRight RestrictAccess -Description "Limit watcher to one mailbox"`

This limits the app to only the mailboxes in that group, even though it holds `Mail.Read` (which normally grants access to all mailboxes).

---

## References

- [Microsoft Graph permissions reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
- [Application Access Policy](https://docs.microsoft.com/en-us/graph/auth-limit-mailbox-access)
- [Client credentials grant](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow)