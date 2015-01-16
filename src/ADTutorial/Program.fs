// Learn more about F# at http://fsharp.net
// See the 'F# Tutorial' project for more help.

[<AutoOpen>]
module AdSettings = 
    open System.Configuration

    let configSetting (key:string) = ConfigurationManager.AppSettings.[key]

    let clientId = configSetting "ad:ClientId"
    let clientSecret = configSetting "ad:ClientSecret"
    let tenantId = configSetting "ad:TenantId"
    let tenantName = configSetting "ad:TenantName"
    let resourceUrl = "https://graph.windows.net"
    let authString = "https://login.windows.net/" + tenantName

[<AutoOpen>]
module AdClient = 
    open Microsoft.Azure.ActiveDirectory.GraphClient
    open Microsoft.IdentityModel.Clients.ActiveDirectory
    open System
    open System.Linq.Expressions
    open System.Threading.Tasks
    open Microsoft.FSharp.Linq.RuntimeHelpers
    
    let getAuthenticationToken (clientId:string) (clientSecret:string) tenantName = 
        let authenticationContext = AuthenticationContext(authString, false)
        let credentials = ClientCredential(clientId, clientSecret)
        let authenticationResult = authenticationContext.AcquireToken(resourceUrl, credentials)
        authenticationResult.AccessToken

    let activeDirectoryClient (tenantId:string) token = 
        let serviceRoot = Uri(resourceUrl + "/" + tenantId)
        let tokenTask = Func<Task<string>>(fun() ->Task.Factory.StartNew<string>(fun() -> token))
        let activeDirectoryClient = ActiveDirectoryClient(serviceRoot, tokenTask)
        activeDirectoryClient

    let client = getAuthenticationToken clientId clientSecret tenantName |> activeDirectoryClient tenantId

    let toExpression<'a> quotationExpression = quotationExpression |> LeafExpressionConverter.QuotationToExpression |> unbox<Expression<'a>>

    let getGroup groupName = 
        let matchExpression = <@Func<IGroup,bool>(fun (group:IGroup) -> group.DisplayName = groupName) @>
        let filter = toExpression<Func<IGroup,bool>> matchExpression
        let groups = client
                        .Groups
                        .Where(filter)
                        .ExecuteAsync()
                        .Result
                        .CurrentPage
                        |> List.ofSeq
        match groups with
        | [] -> None
        | x::[] -> Some (x :?> Group)
        | _ -> raise (Exception("more than one group exists"))

    let addGroup groupName = 
        match getGroup groupName with
        | None ->
            let group = Group()
            group.DisplayName <- groupName
            group.Description <- groupName
            group.MailNickname <- groupName
            group.MailEnabled <- Nullable(false)
            group.SecurityEnabled <- Nullable(true)
            client.Groups.AddGroupAsync(group).Wait()
            Some group
        | Some x -> Some (x :?> Group)

    let getUser userName = 
        let matchExpression = <@Func<IUser,bool>(fun (user:IUser) -> user.DisplayName = userName) @>
        let filter = toExpression<Func<IUser,bool>> matchExpression
        let users = client
                        .Users
                        .Where(filter)
                        .ExecuteAsync()
                        .Result
                        .CurrentPage
                        |> List.ofSeq
        match users with
        | [] -> None
        | x::[] -> Some x
        | _ -> raise (Exception("more than one user exists with that name"))

    let addUser userName = 
        match getUser userName with
        | None ->
            let passwordProfile() =
                let passwd = PasswordProfile()
                passwd.ForceChangePasswordNextLogin <- Nullable(true)
                passwd.Password <- "Ch@ng3NoW!"
                passwd
            let user = User()
            user.PasswordProfile <- passwordProfile()
            user.DisplayName <- userName
            user.UserPrincipalName <- userName + "@fsharptest.onmicrosoft.com"
            user.AccountEnabled <- Nullable(true)
            user.MailNickname <- userName
            client.Users.AddUserAsync(user).Wait()
            Some user
        | Some x -> Some (x :?> User)

    let getMembers (group:Group) = 
        let groupFetcher = (group :> IGroupFetcher)
        let members = groupFetcher.Members.ExecuteAsync().Result
        members.CurrentPage

    let groupContainsUser (group:Group) (user:User) = 
        group |> getMembers |> Seq.map (fun o -> (o :?> User).DisplayName) |> Seq.exists (fun s -> s = user.DisplayName)

    let addUserToGroup (group:Group) (user:User) = 
        match groupContainsUser (group:Group) (user:User) with
        | false ->
            group.Members.Add(user)
            group.UpdateAsync().Wait()
            group
        | true ->
            group

open System
open Microsoft.Azure.ActiveDirectory.GraphClient
[<EntryPoint>]
let main argv = 
    try
        let group = addGroup "newgroup" |> Option.get
        let user = addUser "charlie" |> Option.get
        user |> addUserToGroup group |> ignore
        let group2 = getGroup "newgroup" |> Option.get

        printfn "Group name: %s" group2.DisplayName
        let membersOfGroup = getMembers group
        let members = membersOfGroup |> Seq.map (fun o -> (o :?> User).DisplayName) |> String.concat ", "
        printfn "Members: %s" members
        Console.ReadLine() |> ignore
    with
    | :? Exception as ex -> printfn "%s" ex.Message
    0 
