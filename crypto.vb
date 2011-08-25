Option Explicit On
Option Strict On

Imports System
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Cryptography.Pkcs
Imports System.Text 

Imports System.Net.ServicePointManager

Public Class Crypto

    Dim ai As AlgorithmIdentifier



    Public Sub Headers()
        Dim msgHeaders As System.Net.ServicePointManager

        ' only for IIS 7
        msgHeaders.CheckCertificateRevocationList = True


    End Sub
    Public Shared Function Sign(ByVal buffer() As Byte, ByVal x509Cert As X509Certificate2) As Byte()

        Dim content As ContentInfo
        Dim signedCMS As SignedCms
        Dim cmsSigner As CmsSigner

        Try

            ' Setup the data to be signed
            content = New ContentInfo(buffer) 
            signedCMS = New SignedCms(content, False)
            cmsSigner = New CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, x509Cert)

            ' Now create the signature, the signer will sign the data 
            signedCMS.ComputeSignature(cmsSigner)
            Return signedCMS.Encode

        Catch ex As Exception

        End Try

    End Function

    Private Sub GetCertificate(ByVal certName As X509Certificate2) 

        Dim certStore As X509Store = Nothing
        Dim certCollection As X509Certificate2Collection = Nothing

        Try
            ' Instansiate the certificate store from Personal Certificate, and the Current User 
            certStore = New X509Store(StoreName.My, StoreLocation.CurrentUser)
            certStore.Open(OpenFlags.ReadOnly)
            certCollection = certStore.Certificates.Find(X509FindType.FindBySubjectName , certName, False)
            certStore.Close()
            If certCollection.Count = 0 Then
                Exit Sub
            End If

        Catch ex As CryptographicException
            Throw ex 
        Catch ex1 As Exception
            Throw ex1
        Finally
            certStore.Close()
        End Try

    End Sub

    Public Shared Function Encrypt(ByVal buffer() As Byte, ByVal certName As X509Certificate2) As Byte() 

        Dim content As ContentInfo
        Dim envelopedCMS As EnvelopedCms
        Dim cmsRecipient As CmsRecipient

        Try
            content = New ContentInfo(buffer)
            envelopedCMS = New EnvelopedCms(content) 
            ' Pass a certificate to cmsrecipient to Encrypt
            cmsRecipient = New CmsRecipient(SubjectIdentifierType.IssuerAndSerialNumber, certName)
            envelopedCMS.Encrypt(cmsRecipient) ' Encrypt 
            Return envelopedCMS.Encode ' After the encrypt perform an encode

        Catch ex As CryptographicException
            Throw ex
        Catch ex1 As Exception
            Throw ex1
        End Try

    End Function

    Public Shared Function Decrypt(ByVal buffer() As Byte) As Byte()

        Dim envelopedCMS As EnvelopedCms

        Try
            envelopedCMS = New EnvelopedCms() 
            envelopedCMS.Decode(buffer) ' Convert the PKCS7 encoded data into raw data
            envelopedCMS.Decrypt() ' This will decrypt the key embedded using the key identified in message
            Return envelopedCMS.ContentInfo.Content ' return the content

        Catch ex As CryptographicException
            Throw ex
        Catch ex1 As Exception
            Throw ex1

        End Try
    End Function 

    Public Shared Function Verify(ByVal buffer() As Byte, ByVal certName As X509Certificate2) As Boolean
        Dim signedCMS As SignedCms
        Dim signTrue As Boolean = False

        Try

            signedCMS = New SignedCms() 
            signedCMS.Decode(buffer)

            signedCMS.CheckSignature(New X509Certificate2Collection(certName), False)
            Return True

        Catch ex As CryptographicException
            Return False 
            Throw ex
        Catch ex1 As Exception
            Throw ex1
        End Try
    End Function
End Class