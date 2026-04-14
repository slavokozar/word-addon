await Office.onReady();

function parseJwt(token) {
  const payload = token.split(".")[1];
  const decoded = JSON.parse(atob(payload.replace(/-/g, "+").replace(/_/g, "/")));
  return decoded;
}

async function verifyCurrentUser() {
  try {
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true
    });

    const claims = parseJwt(token);

    const user = {
      name: claims.name,
      email: claims.preferred_username || claims.upn || claims.email,
      oid: claims.oid,
      sub: claims.sub
    };

    console.log("Current user:", user);
    return user;
  } catch (err) {
    console.error("SSO failed", err);
    throw err;
  }
}