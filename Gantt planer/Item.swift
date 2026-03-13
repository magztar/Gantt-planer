//
//  Item.swift
//  Gantt planer
//
//  Created by Magnus Larsson on 2026-03-13.
//

import Foundation
import SwiftData

@Model
final class Item {
    var timestamp: Date
    
    init(timestamp: Date) {
        self.timestamp = timestamp
    }
}
