package dev.erst.gridgrind.excel;

/** Transformer factory that deterministically fails for custom-XML error-translation tests. */
public final class ThrowingTransformerFactory extends javax.xml.transform.TransformerFactory {
  @Override
  public javax.xml.transform.Transformer newTransformer(javax.xml.transform.Source source)
      throws javax.xml.transform.TransformerConfigurationException {
    throw new javax.xml.transform.TransformerConfigurationException("factory boom");
  }

  @Override
  public javax.xml.transform.Transformer newTransformer()
      throws javax.xml.transform.TransformerConfigurationException {
    throw new javax.xml.transform.TransformerConfigurationException("factory boom");
  }

  @Override
  public javax.xml.transform.Templates newTemplates(javax.xml.transform.Source source)
      throws javax.xml.transform.TransformerConfigurationException {
    throw new javax.xml.transform.TransformerConfigurationException("factory boom");
  }

  @Override
  public javax.xml.transform.Source getAssociatedStylesheet(
      javax.xml.transform.Source source, String media, String title, String charset)
      throws javax.xml.transform.TransformerConfigurationException {
    throw new javax.xml.transform.TransformerConfigurationException("factory boom");
  }

  @Override
  public void setURIResolver(javax.xml.transform.URIResolver resolver) {}

  @Override
  public javax.xml.transform.URIResolver getURIResolver() {
    return null;
  }

  @Override
  public void setFeature(String name, boolean value) {}

  @Override
  public boolean getFeature(String name) {
    return false;
  }

  @Override
  public void setAttribute(String name, Object value) {}

  @Override
  public Object getAttribute(String name) {
    return null;
  }

  @Override
  public void setErrorListener(javax.xml.transform.ErrorListener listener) {}

  @Override
  public javax.xml.transform.ErrorListener getErrorListener() {
    return null;
  }
}
