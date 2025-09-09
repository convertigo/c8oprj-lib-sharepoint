// ---- Imports
var I = new JavaImporter(
  Packages.org.apache.hc.client5.http.classic.methods,
  Packages.org.apache.hc.client5.http.impl.classic,
  Packages.org.apache.hc.client5.http.auth,
  Packages.org.apache.hc.client5.http.config,
  Packages.org.apache.hc.client5.http.impl.auth,
  Packages.org.apache.hc.client5.http.protocol,
  Packages.org.apache.hc.client5.http,
  Packages.org.apache.hc.client5.http.cookie,
  Packages.org.apache.hc.client5.http.cookie.BasicCookieStore,
  Packages.org.apache.hc.client5.http.impl.cookie,
  Packages.org.apache.hc.client5.http.impl.routing,
  Packages.org.apache.hc.client5.http.impl.io,
  Packages.org.apache.hc.core5.http,
  Packages.org.apache.hc.core5.http.io.entity,
  Packages.org.apache.hc.core5.http.message,
  Packages.org.apache.hc.core5.util
);

with (I) {
	function Authenticate() {
		// Prefer NTLM (and allow SPNEGO if server offers it)
		let requestConfig = RequestConfig.custom()
		    .setTargetPreferredAuthSchemes(java.util.Arrays.asList(
		        org.apache.hc.client5.http.auth.StandardAuthScheme.NTLM,
		        org.apache.hc.client5.http.auth.StandardAuthScheme.SPNEGO
		    ))
		    .build();
	
		// Credentials provider for NTLM
		let creds = new BasicCredentialsProvider();
		// Workstation is optional; you can pass null. Domain in the user string also works: "DOMAIN\\user"
	
		let jCharArray = new java.lang.String(password).toCharArray();
	
		creds.setCredentials(
		    new AuthScope(null, -1),
		    new NTCredentials(username, jCharArray, null, domain)
		);
	
		let client = HttpClients.custom()
		    .setDefaultCredentialsProvider(creds)
		    .setDefaultRequestConfig(requestConfig)
		  /*
			  .addRequestInterceptorLast(function (request, entity, context) {
		        log.warn(">> " + request.getMethod() + " " + request.getRequestUri());
		        var it = request.headerIterator();
		        while (it.hasNext()) log.warn(">> " + it.next());
		      })
		      .addResponseInterceptorLast(function (response, entity, context) {
		        log.warn("<< " + response.getCode() + " " + response.getReasonPhrase());
		        var it = response.headerIterator();
		        while (it.hasNext()) log.warn("<< " + it.next());
		      })
		  */	
		  .build();
		  return client;
  }
}
