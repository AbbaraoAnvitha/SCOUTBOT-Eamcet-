using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json;
using Microsoft.Rest;
using Microsoft.Rest.Serialization;
using Microsoft.Bot.Connector;

namespace Scout
{
    public partial class HeroCard
    {
        public HeroCard() { }

       public HeroCard(string title = default(string), string subtitle = default(string), string text = default(string), IList<CardImage> images = default(IList<CardImage>), IList<CardAction> buttons = default(IList<CardAction>), CardAction tap = default(CardAction))
        {
                  Title = title;
                  Subtitle = subtitle;
                  Text = text;
                  Images = images;
                  Buttons = buttons;
                  Tap = tap;
                     }
        [JsonProperty(PropertyName = "title")]
          public string Title { get; set; }
   
            [JsonProperty(PropertyName = "subtitle")]
            public string Subtitle { get; set; }
    
            [JsonProperty(PropertyName = "text")]
            public string Text { get; set; }
    
            [JsonProperty(PropertyName = "images")]
            public IList<CardImage> Images { get; set; }
    
            [JsonProperty(PropertyName = "buttons")]
            public IList<CardAction> Buttons { get; set; }
    
            [JsonProperty(PropertyName = "tap")]
            public CardAction Tap { get; set; }
    

    }
}