#!/bin/bash

# <xbar.title>Endpoint Advisor</xbar.title>
# <xbar.version>v1.0</xbar.version>
# <xbar.author>Your Organization</xbar.author>
# <xbar.author.github>EndpointEngineering</xbar.author.github>
# <xbar.desc>Displays endpoint announcements and support info</xbar.desc>
# <xbar.dependencies>jq,curl</xbar.dependencies>

# Configuration
JSON_URL="https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/ContentData2.json"
CACHE_FILE="$HOME/.endpoint-advisor-cache.json"
CACHE_MAX_AGE=900  # 15 minutes in seconds

# Function: Download and cache JSON
fetch_json() {
    if [ ! -f "$CACHE_FILE" ] || [ $(( $(date +%s) - $(stat -f %m "$CACHE_FILE") )) -gt $CACHE_MAX_AGE ]; then
        curl -s "$JSON_URL" -o "$CACHE_FILE" 2>/dev/null
    fi
    
    if [ ! -f "$CACHE_FILE" ]; then
        echo "⚠️ Endpoint Advisor"
        echo "---"
        echo "Failed to load data | color=red"
        echo "Check network connection"
        exit 0
    fi
}

# Function: Strip markdown formatting for display
strip_markdown() {
    # Remove bold **text**
    local text="$1"
    text=$(echo "$text" | sed -E 's/\*\*([^*]+)\*\*/\1/g')
    # Remove italic *text*
    text=$(echo "$text" | sed -E 's/\*([^*]+)\*/\1/g')
    # Remove underline __text__
    text=$(echo "$text" | sed -E 's/__([^_]+)__/\1/g')
    # Remove color tags [color]text[/color]
    text=$(echo "$text" | sed -E 's/\[(\w+)\]([^\[]+)\[\/\1\]/\2/g')
    echo "$text"
}

# Fetch data
fetch_json

# Parse JSON using jq
ANNOUNCEMENT_TEXT=$(jq -r '.Dashboard.Announcements.Default.Text // "No announcements"' "$CACHE_FILE")
ANNOUNCEMENT_DETAILS=$(jq -r '.Dashboard.Announcements.Default.Details // ""' "$CACHE_FILE")
SUPPORT_TEXT=$(jq -r '.Dashboard.Support.Text // "No support info"' "$CACHE_FILE")

# Check for alerts (you can customize this logic)
HAS_ALERT=false

# Menu bar icon and title
if [ "$HAS_ALERT" = true ]; then
    echo "🔴 Endpoint Advisor"
else
    echo "💻 Endpoint Advisor"
fi

echo "---"

# Announcements Section
echo "📢 Announcements | size=14"
CLEAN_ANNOUNCEMENT=$(strip_markdown "$ANNOUNCEMENT_TEXT")
echo "$CLEAN_ANNOUNCEMENT | size=12 trim=false"

if [ -n "$ANNOUNCEMENT_DETAILS" ]; then
    CLEAN_DETAILS=$(strip_markdown "$ANNOUNCEMENT_DETAILS")
    echo "$CLEAN_DETAILS | size=11 color=gray trim=false"
fi

# Announcement Links
LINK_COUNT=$(jq -r '.Dashboard.Announcements.Default.Links | length' "$CACHE_FILE")
if [ "$LINK_COUNT" -gt 0 ]; then
    echo "---"
    for ((i=0; i<$LINK_COUNT; i++)); do
        LINK_NAME=$(jq -r ".Dashboard.Announcements.Default.Links[$i].Name" "$CACHE_FILE")
        LINK_URL=$(jq -r ".Dashboard.Announcements.Default.Links[$i].Url" "$CACHE_FILE")
        echo "🔗 $LINK_NAME | href=$LINK_URL size=11"
    done
fi

echo "---"

# Support Section
echo "🆘 Support | size=14"
CLEAN_SUPPORT=$(strip_markdown "$SUPPORT_TEXT")
echo "$CLEAN_SUPPORT | size=12 trim=false"

# Support Links
SUPPORT_LINK_COUNT=$(jq -r '.Dashboard.Support.Links | length' "$CACHE_FILE")
if [ "$SUPPORT_LINK_COUNT" -gt 0 ]; then
    for ((i=0; i<$SUPPORT_LINK_COUNT; i++)); do
        LINK_NAME=$(jq -r ".Dashboard.Support.Links[$i].Name" "$CACHE_FILE")
        LINK_URL=$(jq -r ".Dashboard.Support.Links[$i].Url" "$CACHE_FILE")
        echo "🔗 $LINK_NAME | href=$LINK_URL size=11"
    done
fi

# Additional Tabs (with platform filtering)
TAB_COUNT=$(jq -r '.AdditionalTabs | length' "$CACHE_FILE")
if [ "$TAB_COUNT" -gt 0 ]; then
    SHOWN_ANY_TAB=false
    
    for ((i=0; i<$TAB_COUNT; i++)); do
        TAB_ENABLED=$(jq -r ".AdditionalTabs[$i].Enabled" "$CACHE_FILE")
        TAB_PLATFORM=$(jq -r ".AdditionalTabs[$i].Platform // \"both\"" "$CACHE_FILE")
        
        # Only show if enabled AND (no platform specified OR platform is macOS)
        if [ "$TAB_ENABLED" = "true" ] && { [ "$TAB_PLATFORM" = "both" ] || [ "$TAB_PLATFORM" = "macOS" ]; }; then
            # Add separator before first tab
            if [ "$SHOWN_ANY_TAB" = false ]; then
                echo "---"
                SHOWN_ANY_TAB=true
            fi
            
            TAB_HEADER=$(jq -r ".AdditionalTabs[$i].TabHeader" "$CACHE_FILE")
            TAB_TEXT=$(jq -r ".AdditionalTabs[$i].Content.Text" "$CACHE_FILE")
            
            echo "📑 $TAB_HEADER | size=14"
            CLEAN_TAB_TEXT=$(strip_markdown "$TAB_TEXT")
            echo "$CLEAN_TAB_TEXT | size=12 trim=false"
            
            # Tab Links
            TAB_LINK_COUNT=$(jq -r ".AdditionalTabs[$i].Content.Links | length" "$CACHE_FILE")
            if [ "$TAB_LINK_COUNT" -gt 0 ]; then
                for ((j=0; j<$TAB_LINK_COUNT; j++)); do
                    LINK_NAME=$(jq -r ".AdditionalTabs[$i].Content.Links[$j].Name" "$CACHE_FILE")
                    LINK_URL=$(jq -r ".AdditionalTabs[$i].Content.Links[$j].Url" "$CACHE_FILE")
                    echo "--🔗 $LINK_NAME | href=$LINK_URL size=11"
                done
            fi
            
            # Separator between tabs
            echo "---"
        fi
    done
fi

# Footer
echo "🔄 Refresh | refresh=true"
echo "ℹ️ About"
echo "--Version: $(jq -r '.version // "1.0.0"' "$CACHE_FILE")"
echo "--Published: $(jq -r '.publishedDate // "Unknown"' "$CACHE_FILE" | cut -d'T' -f1)"
echo "--Cache: $CACHE_FILE | size=10 color=gray"
