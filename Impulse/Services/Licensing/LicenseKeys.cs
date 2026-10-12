namespace Impulse.Services.Licensing
{
    public static class LicenseKeys
    {
        /// <summary>
        /// Public RSA key embedded in the application to verify signed license payloads.
        /// Cannot be used to generate or forge licenses.
        /// </summary>
        public const string PublicKeyXml = @"<RSAKeyValue><Modulus>l2Xh+9uD7cPPdBo90dr4ZVZoxxrORs4P06n5Ap7CwMwPI3JckSrZYcj9kod1B/q2+QJkByhd6Mnjvm02nqyyoEigrAWuzJfEOwheyUVD8E7MOFh91vo1KNUfvRi4liJ4/bqk89XkBNhzY+nWqS1FEC8k/pj5UAp6x28xfpkCa5hcjQq5U3/vuWr6e/dWl55e2P5VoqQaCucr5jbKfcCCHbFn0TKvCNDZFFNV66bjm+8ta3BFWLgJSo+QjuNd8e1AbDc4b2ht1OPduWWfn/Wh1j80DU9TOwLa7WrpPq6hZ5fn5RD0S+gSkrF/+zcZR6wa/SRKchqMlqRH+SecbFoaEQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";
    }
}
