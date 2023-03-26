package main

import (
    "fmt"
    "github.com/emersion/go-imap"
    "github.com/emersion/go-imap/client"
    "github.com/emersion/go-message/mail"
    "net/url"
    "os"
    "strings"
)

func main() {
    // Connect to server
    c, err := client.DialTLS("outlook.office365.com:993", nil)
    if err != nil {
        fmt.Println(err)
        return
    }
    defer c.Logout()

    // Login
    if err := c.Login("your-email@example.com", "your-password"); err != nil {
        fmt.Println(err)
        return
    }

    // Select mailbox
    mbox, err := c.Select("INBOX", false)
    if err != nil {
        fmt.Println(err)
        return
    }

    // Search for all messages
    from := uint32(1)
    to := mbox.Messages
    if to == 0 {
        fmt.Println("No messages in mailbox")
        return
    }
    seqset := new(imap.SeqSet)
    seqset.AddRange(from, to)
    messages := make(chan *imap.Message, 10)
    go func() {
        if err := c.Fetch(seqset, []imap.FetchItem{imap.FetchEnvelope}, messages); err != nil {
            fmt.Println(err)
            return
        }
    }()

    // Compile hyperlinks
    var links []string
    for msg := range messages {
        r := msg.GetBody(&imap.BodySectionName{})
        if r == nil {
            fmt.Println("Server didn't returned message body")
            return
        }
        mr, err := mail.CreateReader(r)
        if err != nil {
            fmt.Println(err)
            return
        }
        for {
            p, err := mr.NextPart()
            if err == nil {
                mediaType, params, _ := p.Header.ContentType()
                if strings.HasPrefix(mediaType, "text/") {
                    body, _ := mail.ReadAll(p)
                    for _, link := range extractLinks(string(body)) {
                        links = append(links, link)
                    }
                }
            } else if err == mail.ErrMissingBoundary {
                break
            } else {
                fmt.Println(err)
                return
            }
        }
    }

    // Remove duplicates
    uniqueLinks := make(map[string]bool)
    for _, link := range links {
        uniqueLinks[link] = true
    }

    // Write output file
    file, err := os.Create("results.txt")
    if err != nil {
        fmt.Println(err)
        return
    }
    defer file.Close()
    for link := range uniqueLinks {
        fmt.Fprintln(file, link)
    }
}

func extractLinks(body string) []string {
    var links []string
    for _, s := range strings.Split(body, " ") {
        u, err := url.Parse(s)
        if err == nil && (u.Scheme == "http" || u.Scheme == "https") {
            links = append(links, s)
        }
    }
    return links
}
